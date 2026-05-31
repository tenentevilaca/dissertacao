"""
App web do Verificador de Citações Acadêmicas.
Execute: python app.py
Acesse:  http://localhost:8000
"""

import asyncio
import json
import os
import shutil
import sys
import tempfile
import threading
import uuid
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse

# Quando empacotado com PyInstaller, recursos ficam em sys._MEIPASS
if getattr(sys, "frozen", False):
    _BASE = Path(sys._MEIPASS)
    sys.path.insert(0, str(_BASE))
else:
    _BASE = Path(__file__).parent

sys.path.insert(0, str(_BASE))

app = FastAPI(title="Verificador de Citações Acadêmicas")

# ── Estado global dos jobs ──────────────────────────────────────────────
_jobs: dict[str, dict] = {}   # job_id → {queue, tmpdir, status, relatorio_dir}


# ══════════════════════════════════════════════════════════════════════════
# ROTAS
# ══════════════════════════════════════════════════════════════════════════

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = _BASE / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/ping")
async def ping():
    return {"ok": True}


@app.get("/manifest.json")
async def manifest():
    return FileResponse(_BASE / "static" / "manifest.json", media_type="application/manifest+json")


@app.get("/icon-{size}.png")
async def icon(size: str):
    return FileResponse(_BASE / "static" / f"icon-{size}.png", media_type="image/png")


@app.post("/iniciar")
async def iniciar_analise(
    dissertacao: UploadFile = File(...),
    referencias: list[UploadFile] = File(...),
    api_key: str = Form(""),
    sem_verificacao: str = Form("false"),
):
    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"citverif_{job_id}_")
    queue: asyncio.Queue = asyncio.Queue()

    _jobs[job_id] = {
        "queue": queue,
        "tmpdir": tmpdir,
        "status": "aguardando",
        "relatorio_dir": None,
    }

    # Salva arquivos enviados
    diss_path = Path(tmpdir) / dissertacao.filename
    diss_path.write_bytes(await dissertacao.read())

    refs_dir = Path(tmpdir) / "referencias"
    refs_dir.mkdir()
    for arq in referencias:
        dest = refs_dir / arq.filename
        dest.write_bytes(await arq.read())

    # Converte sem_verificacao (vem como string do FormData JS)
    sem_verif_bool = str(sem_verificacao).lower() in ("true", "1", "yes")

    # Dispara análise em thread separada (código síncrono)
    loop = asyncio.get_event_loop()
    threading.Thread(
        target=_rodar_analise,
        args=(job_id, str(diss_path), str(refs_dir), api_key, sem_verif_bool, loop),
        daemon=True,
    ).start()

    return {"job_id": job_id}


@app.get("/progresso/{job_id}")
async def progresso(job_id: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")

    async def _stream():
        queue = _jobs[job_id]["queue"]
        while True:
            msg = await queue.get()
            yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
            if msg.get("tipo") in ("concluido", "erro"):
                break

    return StreamingResponse(
        _stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/relatorio/{job_id}/{formato}")
async def baixar_relatorio(job_id: str, formato: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    rel_dir = _jobs[job_id].get("relatorio_dir")
    if not rel_dir:
        raise HTTPException(400, "Relatório ainda não gerado")

    nomes = {"html": "relatorio.html", "txt": "relatorio.txt", "json": "relatorio.json"}
    if formato not in nomes:
        raise HTTPException(400, "Formato inválido")

    caminho = Path(rel_dir) / nomes[formato]
    if not caminho.exists():
        raise HTTPException(404, "Arquivo não encontrado")

    return FileResponse(
        caminho,
        filename=nomes[formato],
        media_type="text/html" if formato == "html" else "application/octet-stream",
    )


@app.delete("/job/{job_id}")
async def limpar_job(job_id: str):
    if job_id in _jobs:
        tmpdir = _jobs[job_id].get("tmpdir")
        if tmpdir and Path(tmpdir).exists():
            shutil.rmtree(tmpdir, ignore_errors=True)
        del _jobs[job_id]
    return {"ok": True}


# ══════════════════════════════════════════════════════════════════════════
# LÓGICA DE ANÁLISE (roda em thread)
# ══════════════════════════════════════════════════════════════════════════

def _enviar(loop: asyncio.AbstractEventLoop, queue: asyncio.Queue, msg: dict):
    asyncio.run_coroutine_threadsafe(queue.put(msg), loop)


def _rodar_analise(
    job_id: str,
    diss_path: str,
    refs_dir: str,
    api_key: str,
    sem_verificacao: bool,
    loop: asyncio.AbstractEventLoop,
):
    queue = _jobs[job_id]["queue"]
    tmpdir = _jobs[job_id]["tmpdir"]
    _jobs[job_id]["status"] = "rodando"

    def log(msg: str, tipo: str = "info"):
        _enviar(loop, queue, {"tipo": tipo, "msg": msg})

    try:
        if not sem_verificacao and not api_key.strip():
            raise ValueError("Chave da API Claude não informada. Informe a chave ou ative 'Apenas cruzamento'.")

        os.environ["ANTHROPIC_API_KEY"] = api_key

        # ── imports locais para não poluir escopo global ──
        import config as cfg
        cfg.DISSERTACAO_PATH    = diss_path
        cfg.REFERENCIAS_FOLDER  = refs_dir
        cfg.ANTHROPIC_API_KEY   = api_key

        from modules.dissertation_parser import DissertationParser
        from modules.reference_reader import ReferenceReader
        from modules.citation_verifier import CitationVerifier, Veredicto
        from modules.report_generator import ReportGenerator

        # ── Etapa 1 ──────────────────────────────────────────────────────
        log("📄 Lendo a dissertação…", "etapa")
        parser = DissertationParser(diss_path)
        parser.carregar()
        citacoes    = parser.extrair_citacoes()
        referencias = parser.extrair_referencias()
        log(f"   {len(citacoes)} citações | {len(referencias)} referências encontradas")

        # ── Etapa 2 ──────────────────────────────────────────────────────
        log("🔗 Cruzando citações com a lista de referências…", "etapa")
        cruzamento       = parser.cruzar_citacoes_referencias()
        citadas_sem_ref  = cruzamento["citadas_sem_referencia"]
        refs_sem_citacao = cruzamento["referenciadas_sem_citacao"]
        pareamentos      = cruzamento["pareamentos"]
        log(f"   {len(pareamentos)} pares | {len(citadas_sem_ref)} sem referência | {len(refs_sem_citacao)} ref. sem citação")

        # ── Etapa 3 ──────────────────────────────────────────────────────
        log("📂 Lendo os documentos da pasta de referências…", "etapa")
        reader = ReferenceReader(refs_dir)

        # Versão com callback de progresso
        from config import EXTENSOES_SUPORTADAS
        arquivos = [
            f for f in Path(refs_dir).rglob("*")
            if f.is_file() and f.suffix.lower() in EXTENSOES_SUPORTADAS
        ]
        log(f"   {len(arquivos)} arquivo(s) encontrado(s)")
        for arq in arquivos:
            doc = reader._ler_arquivo(arq)
            if doc:
                reader.documentos.append(doc)
                log(f"   ✓ {arq.name}")
        reader._construir_indice()

        # ── Etapa 4 ──────────────────────────────────────────────────────
        verificacoes = []
        if not sem_verificacao and pareamentos:
            log(f"🤖 Verificando {len(pareamentos)} citações com IA…", "etapa")
            verifier = CitationVerifier(reader)
            for i, (cit, ref) in enumerate(pareamentos, 1):
                log(f"   [{i}/{len(pareamentos)}] {cit.texto_original[:70]}…")
                resultado = verifier.verificar(cit, ref)
                emoji = {"CORRETO": "✓", "INCORRETO": "✗", "PARCIAL": "⚠", "SEM_FONTE": "?", "ERRO": "!"}.get(resultado.veredicto.value, "?")
                log(f"   {emoji} {resultado.veredicto.value} — {resultado.justificativa[:80]}")
                verificacoes.append(resultado)
        else:
            log("⏭ Verificação semântica ignorada.", "etapa")

        # ── Relatório ─────────────────────────────────────────────────────
        log("📊 Gerando relatórios…", "etapa")
        rel_dir = Path(tmpdir) / "relatorio"
        gen = ReportGenerator(
            citadas_sem_ref=citadas_sem_ref,
            refs_sem_citacao=refs_sem_citacao,
            verificacoes=verificacoes,
            caminho_dissertacao=str(diss_path),
        )
        gen.gerar_txt(rel_dir / "relatorio.txt")
        gen.gerar_html(rel_dir / "relatorio.html")
        gen.gerar_json(rel_dir / "relatorio.json")

        _jobs[job_id]["relatorio_dir"] = str(rel_dir)
        _jobs[job_id]["status"] = "concluido"

        # Resumo final
        corretas   = sum(1 for v in verificacoes if v.veredicto.value == "CORRETO")
        incorretas = sum(1 for v in verificacoes if v.veredicto.value == "INCORRETO")
        parciais   = sum(1 for v in verificacoes if v.veredicto.value == "PARCIAL")
        sem_fonte  = sum(1 for v in verificacoes if v.veredicto.value == "SEM_FONTE")

        _enviar(loop, queue, {
            "tipo": "concluido",
            "msg": "Análise concluída!",
            "resumo": {
                "corretas":          corretas,
                "incorretas":        incorretas,
                "parciais":          parciais,
                "sem_fonte":         sem_fonte,
                "citadas_sem_ref":   len(citadas_sem_ref),
                "refs_sem_citacao":  len(refs_sem_citacao),
                "total_citacoes":    len(citacoes),
                "total_referencias": len(referencias),
            },
        })

    except Exception as exc:
        import traceback
        _jobs[job_id]["status"] = "erro"
        _enviar(loop, queue, {"tipo": "erro", "msg": f"Erro: {exc}", "detalhe": traceback.format_exc()})


# ══════════════════════════════════════════════════════════════════════════
# ENTRYPOINT
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    print("\n" + "═" * 60)
    print("  VERIFICADOR DE CITAÇÕES ACADÊMICAS — APP WEB")
    print("  Acesse no navegador:  http://localhost:8000")
    print("  Na mesma rede/tablet: http://<SEU-IP>:8000")
    print("═" * 60 + "\n")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
