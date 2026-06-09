"""
Analisador de Dissertações Acadêmicas
Execute: python app.py
Acesse:  http://localhost:8001
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

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse

if getattr(sys, "frozen", False):
    _BASE = Path(sys._MEIPASS)
    sys.path.insert(0, str(_BASE))
else:
    _BASE = Path(__file__).parent

sys.path.insert(0, str(_BASE))

app = FastAPI(title="Analisador de Dissertações Acadêmicas")

_jobs: dict[str, dict] = {}


# ══════════════════════════════════════════════════════════════════════════
# ROTAS ESTÁTICAS
# ══════════════════════════════════════════════════════════════════════════

@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse((_BASE / "static" / "index.html").read_text(encoding="utf-8"))


@app.get("/ping")
async def ping():
    return {"ok": True}


# ══════════════════════════════════════════════════════════════════════════
# INICIAR ANÁLISE
# ══════════════════════════════════════════════════════════════════════════

@app.post("/iniciar")
async def iniciar_analise(
    dissertacao: UploadFile = File(...),
    apoio: UploadFile = File(None),
    api_key: str = Form(""),
):
    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"analisador_{job_id}_")
    queue: asyncio.Queue = asyncio.Queue()

    _jobs[job_id] = {
        "queue": queue,
        "tmpdir": tmpdir,
        "status": "aguardando",
        "saida_dir": None,
    }

    # Salva arquivos enviados
    diss_path = Path(tmpdir) / (dissertacao.filename or "dissertacao.docx")
    with open(diss_path, "wb") as f:
        f.write(await dissertacao.read())

    apoio_path = None
    if apoio and apoio.filename:
        apoio_path = Path(tmpdir) / (apoio.filename or "apoio.zip")
        with open(apoio_path, "wb") as f:
            f.write(await apoio.read())

    if api_key:
        os.environ["ANTHROPIC_API_KEY"] = api_key

    # Dispara processamento em thread separada
    thread = threading.Thread(
        target=_executar_analise,
        args=(job_id, diss_path, apoio_path, queue),
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id}


# ══════════════════════════════════════════════════════════════════════════
# STREAM DE PROGRESSO (SSE)
# ══════════════════════════════════════════════════════════════════════════

@app.get("/progresso/{job_id}")
async def progresso(job_id: str):
    if job_id not in _jobs:
        return HTMLResponse("Job não encontrado", status_code=404)

    async def gerar():
        queue = _jobs[job_id]["queue"]
        while True:
            try:
                msg = await asyncio.wait_for(queue.get(), timeout=120)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg.get("tipo") in ("concluido", "erro"):
                    break
            except asyncio.TimeoutError:
                yield "data: {\"tipo\": \"ping\"}\n\n"

    return StreamingResponse(gerar(), media_type="text/event-stream")


# ══════════════════════════════════════════════════════════════════════════
# DOWNLOAD DOS RESULTADOS
# ══════════════════════════════════════════════════════════════════════════

@app.get("/download/{job_id}/{arquivo}")
async def download(job_id: str, arquivo: str):
    if job_id not in _jobs:
        return HTMLResponse("Job não encontrado", status_code=404)

    saida_dir = _jobs[job_id].get("saida_dir")
    if not saida_dir:
        return HTMLResponse("Resultados ainda não disponíveis", status_code=404)

    # Segurança: apenas nomes simples, sem path traversal
    if "/" in arquivo or "\\" in arquivo or ".." in arquivo:
        return HTMLResponse("Arquivo inválido", status_code=400)

    caminho = Path(saida_dir) / arquivo
    if not caminho.exists():
        return HTMLResponse("Arquivo não encontrado", status_code=404)

    media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(str(caminho), media_type=media, filename=arquivo)


@app.get("/metadados/{job_id}")
async def metadados(job_id: str):
    if job_id not in _jobs:
        return {"erro": "Job não encontrado"}
    return _jobs[job_id].get("metadados", {})


# ══════════════════════════════════════════════════════════════════════════
# PROCESSAMENTO EM BACKGROUND
# ══════════════════════════════════════════════════════════════════════════

def _pub(queue: asyncio.Queue, loop: asyncio.AbstractEventLoop, tipo: str, mensagem: str, **extra):
    dados = {"tipo": tipo, "mensagem": mensagem, **extra}
    asyncio.run_coroutine_threadsafe(queue.put(dados), loop)


def _executar_analise(
    job_id: str,
    diss_path: Path,
    apoio_path: Path | None,
    queue: asyncio.Queue,
):
    loop = asyncio.new_event_loop()

    def pub(tipo: str, msg: str, **kw):
        _pub(queue, loop, tipo, msg, **kw)

    loop.run_in_executor = None  # não usar

    try:
        from config import ANTHROPIC_API_KEY as _KEY
        if not _KEY:
            raise ValueError(
                "Chave da API não configurada. Informe a chave no formulário ou "
                "crie um arquivo .env com ANTHROPIC_API_KEY=..."
            )

        from modules.metadata_extractor import extrair_metadados
        from modules.chapter_splitter import dividir_em_capitulos
        from modules.chapter_analyzer import analisar_capitulo
        from modules.support_searcher import extrair_materiais_apoio, buscar_apoio_para_capitulo
        from modules.document_rebuilder import gerar_relatorio, gerar_dissertacao_reescrita

        tmpdir = Path(_jobs[job_id]["tmpdir"])
        saida_dir = tmpdir / "saida"
        saida_dir.mkdir(exist_ok=True)
        _jobs[job_id]["saida_dir"] = str(saida_dir)

        # ── 1. Metadados ──────────────────────────────────────────────
        pub("progresso", "Extraindo metadados da dissertação...", pct=5)
        metadados = extrair_metadados(diss_path)
        _jobs[job_id]["metadados"] = metadados

        autor = metadados.get("autor", "Autor não identificado")
        titulo = metadados.get("titulo", "Dissertação")
        pub("progresso", f"✓ Metadados: {autor} — {titulo}", pct=10)

        # ── 2. Dividir em capítulos ───────────────────────────────────
        pub("progresso", "Dividindo dissertação em capítulos...", pct=12)
        capitulos = dividir_em_capitulos(diss_path)
        pub("progresso", f"✓ {len(capitulos)} capítulos identificados", pct=15)

        # ── 3. Material de apoio ──────────────────────────────────────
        materiais = []
        if apoio_path:
            pub("progresso", "Descompactando e indexando material de apoio...", pct=17)
            pasta_apoio = tmpdir / "apoio"
            try:
                materiais = extrair_materiais_apoio(apoio_path, pasta_apoio)
                pub("progresso", f"✓ {len(materiais)} arquivos de apoio carregados", pct=20)
            except Exception as e:
                pub("aviso", f"Material de apoio: {e} — prosseguindo sem ele.")

        # ── 4. Análise + reescrita por capítulo ───────────────────────
        resultados: list[dict] = []
        total = len([c for c in capitulos if c["numero"] > 0])
        feitos = 0

        for cap in capitulos:
            if cap["numero"] == 0:
                # Pré-textual: não analisa, copia original
                resultados.append({
                    "numero": 0,
                    "titulo": cap["titulo"],
                    "texto_original": cap["texto"],
                    "texto_reescrito": cap["texto"],
                    "diagnostico": {},
                })
                continue

            pct_base = 20 + int(feitos / total * 65)
            pub("progresso", f"Analisando capítulo {cap['numero']}: {cap['titulo']}...", pct=pct_base)

            # Busca material de apoio relevante para este capítulo
            trechos_apoio = []
            if materiais:
                try:
                    trechos_apoio = buscar_apoio_para_capitulo(
                        cap,
                        materiais,
                        on_progress=lambda m: pub("progresso", m, pct=pct_base + 1),
                    )
                    if trechos_apoio:
                        pub("progresso", f"  ✓ {len(trechos_apoio)} trechos de apoio encontrados", pct=pct_base + 2)
                except Exception as e:
                    pub("aviso", f"  ⚠ Apoio para cap {cap['numero']}: {e}")

            # Análise crítica + reescrita
            try:
                resultado = analisar_capitulo(
                    cap,
                    metadados,
                    trechos_apoio,
                    on_progress=lambda m: pub("progresso", m, pct=pct_base + 3),
                )
                resultado["trechos_apoio"] = trechos_apoio
                resultados.append(resultado)
                pontuacao = resultado["diagnostico"].get("pontuacao_geral", "?")
                pub("progresso", f"  ✓ Capítulo {cap['numero']} concluído (nota: {pontuacao}/10)", pct=pct_base + 5)
            except Exception as e:
                pub("aviso", f"  ⚠ Erro no capítulo {cap['numero']}: {e}")
                resultados.append({
                    "numero": cap["numero"],
                    "titulo": cap["titulo"],
                    "texto_original": cap["texto"],
                    "texto_reescrito": cap["texto"],
                    "diagnostico": {"erro": str(e)},
                    "trechos_apoio": [],
                })

            feitos += 1

        # ── 5. Gerar relatório ────────────────────────────────────────
        pub("progresso", "Gerando relatório de análise...", pct=87)
        try:
            caminho_relatorio = gerar_relatorio(metadados, resultados, saida_dir)
            pub("progresso", f"✓ Relatório salvo: {caminho_relatorio.name}", pct=90)
        except Exception as e:
            pub("aviso", f"Erro ao gerar relatório: {e}")

        # ── 6. Gerar dissertação reescrita ────────────────────────────
        pub("progresso", "Gerando dissertação reescrita...", pct=92)
        try:
            caminho_diss = gerar_dissertacao_reescrita(
                diss_path, metadados, resultados, saida_dir
            )
            pub("progresso", f"✓ Dissertação reescrita salva: {caminho_diss.name}", pct=97)
        except Exception as e:
            pub("aviso", f"Erro ao gerar dissertação reescrita: {e}")

        # ── 7. Concluído ──────────────────────────────────────────────
        pub(
            "concluido",
            "Análise concluída com sucesso!",
            pct=100,
            arquivos=["relatorio_analise.docx", "dissertacao_reescrita.docx"],
            metadados=metadados,
        )

    except Exception as e:
        import traceback
        pub("erro", f"Erro crítico: {e}\n{traceback.format_exc()}", pct=0)
    finally:
        loop.close()


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    print("=" * 55)
    print("  ANALISADOR DE DISSERTAÇÕES ACADÊMICAS")
    print("  Acesse: http://localhost:8001")
    print("=" * 55)
    uvicorn.run(app, host="0.0.0.0", port=8001)
