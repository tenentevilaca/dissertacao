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

app = FastAPI(title="Ferramentas Acadêmicas")

# Limite de upload: 1 GB (suporte a pastas de material de apoio grandes)
_1GB = 1_073_741_824
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.datastructures import Headers

class LargeUploadMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        request._max_content_size = _1GB
        return await call_next(request)

app.add_middleware(LargeUploadMiddleware)

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


@app.post("/iniciar-analise-dissertacao")
async def iniciar_analise_dissertacao(
    dissertacao: UploadFile = File(...),
    dissertacao_drive_url: str = Form(""),
    material_apoio: Optional[UploadFile] = File(None),
    material_drive_url: str = Form(""),
    api_key: str = Form(""),
):
    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"analise_{job_id}_")
    queue: asyncio.Queue = asyncio.Queue()

    _jobs[job_id] = {
        "queue": queue,
        "tmpdir": tmpdir,
        "status": "aguardando",
        "relatorio_dir": None,
        "tipo": "analise",
    }

    diss_bytes = await dissertacao.read()
    diss_path = Path(tmpdir) / dissertacao.filename
    diss_path.write_bytes(diss_bytes)

    diss_drive = dissertacao_drive_url.strip() if dissertacao_drive_url else ""

    material_path: Optional[str] = None
    if material_apoio and material_apoio.filename:
        mat_path = Path(tmpdir) / material_apoio.filename
        mat_path.write_bytes(await material_apoio.read())
        material_path = str(mat_path)

    drive_url = material_drive_url.strip() if material_drive_url else ""

    loop = asyncio.get_event_loop()
    threading.Thread(
        target=_rodar_analise_dissertacao,
        args=(job_id, str(diss_path), diss_drive, material_path, drive_url, api_key, loop),
        daemon=True,
    ).start()

    return {"job_id": job_id}


@app.get("/progresso-analise/{job_id}")
async def progresso_analise(job_id: str):
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
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/relatorio-analise/{job_id}/{formato}")
async def baixar_relatorio_analise(job_id: str, formato: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    rel_dir = _jobs[job_id].get("relatorio_dir")
    if not rel_dir:
        raise HTTPException(400, "Relatório ainda não gerado")

    nomes = {
        "html": "relatorio_analise.html",
        "json": "relatorio_analise.json",
        "txt":  "relatorio_analise.txt",
    }
    if formato not in nomes:
        raise HTTPException(400, "Formato inválido")

    caminho = Path(rel_dir) / nomes[formato]
    if not caminho.exists():
        raise HTTPException(404, "Arquivo não encontrado")

    media = "text/html" if formato == "html" else "application/octet-stream"
    return FileResponse(caminho, filename=nomes[formato], media_type=media)


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


def _rodar_analise_dissertacao(
    job_id: str,
    diss_path: str,
    diss_drive_url: str,
    material_path: Optional[str],
    drive_url: str,
    api_key: str,
    loop: asyncio.AbstractEventLoop,
):
    queue = _jobs[job_id]["queue"]
    tmpdir = _jobs[job_id]["tmpdir"]
    _jobs[job_id]["status"] = "rodando"

    def log(msg: str, tipo: str = "info"):
        _enviar(loop, queue, {"tipo": tipo, "msg": msg})

    try:
        if api_key.strip():
            os.environ["ANTHROPIC_API_KEY"] = api_key

        from analisador import (
            analisar_dissertacao,
            extrair_material_de_zip,
            carregar_material_de_pasta,
            gerar_relatorio_analise,
            baixar_google_drive,
            baixar_google_drive_bytes,
            _ler_bytes as _ler_bytes_apoio,
        )

        # Baixa dissertação do Drive se necessário
        diss_final = Path(diss_path)
        if diss_drive_url:
            log("☁️ Baixando dissertação do Google Drive…", "etapa")
            try:
                diss_final = baixar_google_drive(
                    diss_drive_url,
                    Path(tmpdir) / "dissertacao_drive",
                    log_fn=log,
                )
                log(f"   ✓ Dissertação baixada: {diss_final.name}")
            except Exception as e:
                log(f"   ✗ Erro ao baixar dissertação do Drive: {e}", "erro")
                raise

        # Carrega material de apoio
        material: dict[str, str] = {}
        destino = Path(tmpdir) / "material_extraido"

        # Prioridade: arquivo enviado > URL do Drive
        if material_path:
            mat = Path(material_path)
            if mat.suffix.lower() == ".zip":
                log("📦 Lendo arquivos do ZIP em memória…", "etapa")
                material = extrair_material_de_zip(mat)
            else:
                from verificador import ler_arquivo
                texto = ler_arquivo(mat)
                if texto and len(texto) > 100:
                    material = {mat.name: texto}
            log(f"   {len(material)} arquivo(s) de apoio carregado(s)")

        elif drive_url:
            log("☁️ Baixando material de apoio do Google Drive…", "etapa")
            try:
                drive_bytes, drive_ext = baixar_google_drive_bytes(drive_url, log_fn=log)
                if drive_ext.lower() == ".zip":
                    log("📦 Lendo arquivos do ZIP em memória…", "etapa")
                    material = extrair_material_de_zip(drive_bytes)
                else:
                    texto = _ler_bytes_apoio(drive_bytes, drive_ext)
                    if texto and len(texto) > 100:
                        material = {f"arquivo_drive{drive_ext}": texto}
                log(f"   {len(material)} arquivo(s) de apoio carregado(s)")
            except Exception as e:
                log(f"   ⚠ Erro ao baixar do Drive: {e}", "warn")
                log("   Continuando sem material de apoio…", "warn")

        resultado = analisar_dissertacao(
            diss_path=str(diss_final),
            material=material,
            api_key=api_key,
            log_fn=log,
        )

        log("📊 Gerando relatório…", "etapa")
        rel_dir = Path(tmpdir) / "relatorio"
        gerar_relatorio_analise(resultado, rel_dir)

        _jobs[job_id]["relatorio_dir"] = str(rel_dir)
        _jobs[job_id]["status"] = "concluido"

        caps = resultado.get("capitulos") or []
        ponts = [c["analise"].get("pontuacao_geral", "") for c in caps]

        _enviar(loop, queue, {
            "tipo": "concluido",
            "msg": "Análise concluída!",
            "resumo": {
                "n_capitulos":   len(caps),
                "aprovados":     sum(1 for p in ponts if p == "APROVADO"),
                "ressalvas":     sum(1 for p in ponts if p == "APROVADO_COM_RESSALVAS"),
                "revisao":       sum(1 for p in ponts if p == "REQUER_REVISAO"),
                "n_apoio":       resultado.get("n_arquivos_apoio", 0),
                "autor":         (resultado.get("metadados") or {}).get("autor") or "—",
                "titulo":        (resultado.get("metadados") or {}).get("titulo") or "—",
            },
        })

    except Exception as exc:
        import traceback
        _jobs[job_id]["status"] = "erro"
        _enviar(loop, queue, {"tipo": "erro", "msg": f"Erro: {exc}", "detalhe": traceback.format_exc()})


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
        if api_key.strip():
            os.environ["ANTHROPIC_API_KEY"] = api_key

        from verificador import analisar, gerar_relatorio

        resultado = analisar(
            diss_path=diss_path,
            refs_dir=refs_dir,
            api_key=api_key,
            sem_verificacao=sem_verificacao,
            log_fn=log,
        )

        log("📊 Gerando relatórios…", "etapa")
        rel_dir = Path(tmpdir) / "relatorio"
        gerar_relatorio(resultado, rel_dir, diss_path)

        _jobs[job_id]["relatorio_dir"] = str(rel_dir)
        _jobs[job_id]["status"] = "concluido"

        verificacoes     = resultado["verificacoes"]
        citadas_sem_ref  = resultado["citadas_sem_ref"]
        refs_sem_citacao = resultado["refs_sem_citacao"]
        citacoes         = resultado["citacoes"]
        referencias      = resultado["referencias"]

        corretas   = sum(1 for v in verificacoes if v.get("veredicto") == "CORRETO")
        incorretas = sum(1 for v in verificacoes if v.get("veredicto") == "INCORRETO")
        parciais   = sum(1 for v in verificacoes if v.get("veredicto") == "PARCIAL")
        sem_fonte  = sum(1 for v in verificacoes if v.get("veredicto") == "SEM_FONTE")

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
    print("  FERRAMENTAS ACADÊMICAS — APP WEB")
    print("  Acesse no navegador:  http://localhost:8000")
    print("  Na mesma rede/tablet: http://<SEU-IP>:8000")
    print("═" * 60 + "\n")
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=False,
        # Limite de corpo HTTP: 1 GB
        h11_max_incomplete_event_size=1_073_741_824,
    )
