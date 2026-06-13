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
    return HTMLResponse(
        html_path.read_text(encoding="utf-8"),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
    )


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
    material_localizado_job_id: str = Form(""),
    api_key: str = Form(""),
):
    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"analise_{job_id}_")
    queue: asyncio.Queue = asyncio.Queue()

    _jobs[job_id] = {
        "queue": queue,
        "logs": [],
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
    loc_job_id = material_localizado_job_id.strip() if material_localizado_job_id else ""

    loop = asyncio.get_event_loop()
    threading.Thread(
        target=_rodar_analise_dissertacao,
        args=(job_id, str(diss_path), diss_drive, material_path, drive_url, loc_job_id, api_key, loop),
        daemon=True,
    ).start()

    return {"job_id": job_id}


@app.get("/status-analise/{job_id}")
async def status_analise(job_id: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    job = _jobs[job_id]
    return {
        "status": job["status"],
        "logs": job["logs"],
        "relatorio_dir": job.get("relatorio_dir"),
    }


@app.get("/progresso-analise/{job_id}")
async def progresso_analise(job_id: str, desde: int = 0):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")

    async def _stream():
        job = _jobs[job_id]
        # Envia logs já acumulados (reconexão)
        for msg in job["logs"][desde:]:
            yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
            if msg.get("tipo") in ("concluido", "erro"):
                return
        # Continua com novos eventos da fila
        queue = job["queue"]
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


@app.get("/dissertacao-revisada/{job_id}")
async def baixar_dissertacao_revisada(job_id: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    docx_path = _jobs[job_id].get("docx_revisado")
    if not docx_path or not Path(docx_path).exists():
        raise HTTPException(404, "Dissertação revisada não disponível")

    meta = _jobs[job_id].get("metadados") or {}
    nome = "dissertacao_revisada.docx"
    return FileResponse(
        docx_path, filename=nome,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/iniciar-localizar-referencias")
async def iniciar_localizar_referencias(
    referencias_texto: str = Form(""),
    referencias_arquivo: Optional[UploadFile] = File(None),
    material_apoio: list[UploadFile] = File(default=[]),
    material_drive_url: str = Form(""),
    api_key: str = Form(""),
):
    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"refs_{job_id}_")

    _jobs[job_id] = {
        "tmpdir": tmpdir,
        "status": "rodando",
        "tipo": "localizar_referencias",
        "material_bytes": {},
        "material_texto": {},
        "arquivos_relevantes": [],
        "resultado": None,
        "logs": [],
    }

    refs_texto_final = referencias_texto or ""
    refs_path: Optional[str] = None
    if referencias_arquivo and referencias_arquivo.filename:
        rp = Path(tmpdir) / referencias_arquivo.filename
        rp.write_bytes(await referencias_arquivo.read())
        refs_path = str(rp)

    mat_paths: list[str] = []
    for arq in material_apoio:
        if arq and arq.filename:
            mp = Path(tmpdir) / arq.filename
            mp.write_bytes(await arq.read())
            mat_paths.append(str(mp))

    drive_url = material_drive_url.strip() if material_drive_url else ""

    threading.Thread(
        target=_rodar_localizar_referencias,
        args=(job_id, refs_texto_final, refs_path, mat_paths, drive_url, api_key.strip()),
        daemon=True,
    ).start()

    return {"job_id": job_id}


@app.get("/status-localizar-referencias/{job_id}")
async def status_localizar_referencias(job_id: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    job = _jobs[job_id]
    return {"status": job["status"], "resultado": job.get("resultado"), "logs": job.get("logs", [])}


@app.get("/material-arquivo/{job_id}/{nome_arquivo}")
async def baixar_material_arquivo(job_id: str, nome_arquivo: str):
    if job_id not in _jobs:
        raise HTTPException(404, "Job não encontrado")
    dados = _jobs[job_id].get("material_bytes", {}).get(nome_arquivo)
    if dados is None:
        raise HTTPException(404, "Arquivo não encontrado")
    from fastapi.responses import Response
    return Response(content=dados, media_type="application/octet-stream", headers={
        "Content-Disposition": f'attachment; filename="{nome_arquivo}"'
    })


def _rodar_localizar_referencias(
    job_id: str,
    refs_texto: str,
    refs_path: Optional[str],
    mat_paths: list[str],
    drive_url: str,
    api_key: str = "",
):
    def log(msg: str):
        _jobs[job_id]["logs"].append(msg)

    try:
        from analisador import (
            parsear_lista_referencias,
            localizar_referencias,
            extrair_material_de_zip_com_bytes,
            baixar_google_drive_bytes,
            baixar_google_drive_path,
            _ler_bytes as _ler_bytes_apoio,
        )
        from verificador import ler_arquivo

        # ── Lista de referências ─────────────────────────────────────────
        log("📑 Lendo lista de referências…")
        texto_refs = refs_texto or ""
        if refs_path:
            p = Path(refs_path)
            if p.suffix.lower() == ".txt":
                texto_refs += "\n" + p.read_text(encoding="utf-8", errors="replace")
            else:
                texto_refs += "\n" + ler_arquivo(p)

        referencias = parsear_lista_referencias(texto_refs)
        log(f"   {len(referencias)} referência(s) identificada(s)")

        # ── Material de apoio ────────────────────────────────────────────
        log("📂 Lendo material de apoio…")
        material: dict[str, str] = {}
        material_bytes: dict[str, bytes] = {}

        for mat_path in mat_paths:
            mp = Path(mat_path)
            if mp.suffix.lower() == ".zip":
                log(f"   Extraindo {mp.name}…")
                for nome, (texto, dados) in extrair_material_de_zip_com_bytes(mp, log).items():
                    material[nome] = texto
                    material_bytes[nome] = dados
                    log(f"   ✓ {nome} ({len(texto):,} chars)")
            else:
                texto = ler_arquivo(mp)
                if texto and len(texto) > 100:
                    nome = mp.name[:80]
                    material[nome] = texto
                    material_bytes[nome] = mp.read_bytes()
                    log(f"   ✓ {nome} ({len(texto):,} chars)")
                else:
                    log(f"   ✗ {mp.name} — não foi possível ler")

        if not mat_paths and drive_url:
            log("   Baixando material do Google Drive…")
            drive_path, drive_ext = baixar_google_drive_path(drive_url, log_fn=log)
            if drive_ext.lower() == ".zip":
                try:
                    for nome, (texto, dados) in extrair_material_de_zip_com_bytes(drive_path, log).items():
                        material[nome] = texto
                        material_bytes[nome] = dados
                        log(f"   ✓ {nome} ({len(texto):,} chars)")
                finally:
                    try:
                        drive_path.unlink(missing_ok=True)
                    except OSError:
                        pass
            else:
                drive_bytes = drive_path.read_bytes()
                try:
                    drive_path.unlink(missing_ok=True)
                except OSError:
                    pass
                texto = _ler_bytes_apoio(drive_bytes, drive_ext)
                if texto and len(texto) > 100:
                    nome = f"arquivo_drive{drive_ext}"
                    material[nome] = texto
                    material_bytes[nome] = drive_bytes
                    log(f"   ✓ {nome} ({len(texto):,} chars)")

        log(f"   {len(material)} arquivo(s) de apoio lido(s)")

        # ── Localização das referências ──────────────────────────────────
        log(f"🔎 Procurando {len(referencias)} referência(s) em {len(material)} arquivo(s)…")

        def log_ref(i: int, total: int, ref: str, encontrado: bool, arquivos: list[str]):
            if encontrado:
                log(f"   [{i}/{total}] ✓ {ref[:70]} → {', '.join(arquivos)}")
            else:
                log(f"   [{i}/{total}] ✗ {ref[:70]} — não encontrado")

        resultado = localizar_referencias(referencias, material, api_key, log_fn=log_ref)

        arquivos_relevantes = sorted({
            nome for item in resultado for nome in item["arquivos"]
        })

        _jobs[job_id]["material_bytes"] = material_bytes
        _jobs[job_id]["material_texto"] = material
        _jobs[job_id]["arquivos_relevantes"] = arquivos_relevantes
        _jobs[job_id]["resultado"] = {
            "itens": resultado,
            "n_referencias": len(referencias),
            "n_arquivos": len(material),
            "n_arquivos_relevantes": len(arquivos_relevantes),
        }
        _jobs[job_id]["status"] = "concluido"

    except Exception as exc:
        import traceback
        _jobs[job_id]["status"] = "erro"
        _jobs[job_id]["resultado"] = {"erro": str(exc), "detalhe": traceback.format_exc()}


@app.post("/iniciar")
async def iniciar_analise(
    dissertacao: Optional[UploadFile] = File(None),
    dissertacao_drive_url: str = Form(""),
    referencias: list[UploadFile] = File(default=[]),
    referencias_drive_url: str = Form(""),
    api_key: str = Form(""),
    sem_verificacao: str = Form("false"),
    material_localizado_job_id: str = Form(""),
):
    from analisador import extrair_material_de_zip_com_bytes, baixar_google_drive_bytes, baixar_google_drive_path

    job_id = str(uuid.uuid4())
    tmpdir = tempfile.mkdtemp(prefix=f"citverif_{job_id}_")
    queue: asyncio.Queue = asyncio.Queue()

    _jobs[job_id] = {
        "queue": queue,
        "tmpdir": tmpdir,
        "status": "aguardando",
        "relatorio_dir": None,
    }

    # Salva a dissertação (arquivo ou link do Google Drive)
    if dissertacao and dissertacao.filename:
        diss_path = Path(tmpdir) / dissertacao.filename
        diss_path.write_bytes(await dissertacao.read())
    elif dissertacao_drive_url.strip():
        diss_src_path, diss_ext = baixar_google_drive_path(dissertacao_drive_url.strip())
        diss_path = Path(tmpdir) / f"dissertacao{diss_ext}"
        import shutil
        shutil.move(str(diss_src_path), str(diss_path))
    else:
        raise HTTPException(400, "Envie a dissertação como arquivo ou link do Google Drive")

    refs_dir = Path(tmpdir) / "referencias"
    refs_dir.mkdir()

    def _salvar_referencia(nome: str, dados, from_path: bool = False):
        suffix = Path(nome).suffix.lower()
        if suffix == ".zip":
            for sub_nome, (_texto, sub_dados) in extrair_material_de_zip_com_bytes(dados, log).items():
                (refs_dir / sub_nome).write_bytes(sub_dados)
            if from_path:
                try:
                    Path(dados).unlink(missing_ok=True)
                except OSError:
                    pass
        else:
            if from_path:
                import shutil
                shutil.move(str(dados), str(refs_dir / nome))
            else:
                (refs_dir / nome).write_bytes(dados)

    for arq in referencias:
        if arq and arq.filename:
            _salvar_referencia(arq.filename, await arq.read())

    if referencias_drive_url.strip():
        drive_path, drive_ext = baixar_google_drive_path(referencias_drive_url.strip())
        _salvar_referencia(f"referencias_drive{drive_ext}", drive_path, from_path=True)

    # Converte sem_verificacao (vem como string do FormData JS)
    sem_verif_bool = str(sem_verificacao).lower() in ("true", "1", "yes")
    loc_job_id = material_localizado_job_id.strip() if material_localizado_job_id else ""

    # Dispara análise em thread separada (código síncrono)
    loop = asyncio.get_event_loop()
    threading.Thread(
        target=_rodar_analise,
        args=(job_id, str(diss_path), str(refs_dir), api_key, sem_verif_bool, loc_job_id, loop),
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
    loc_job_id: str,
    api_key: str,
    loop: asyncio.AbstractEventLoop,
):
    queue = _jobs[job_id]["queue"]
    tmpdir = _jobs[job_id]["tmpdir"]
    _jobs[job_id]["status"] = "rodando"

    def log(msg: str, tipo: str = "info"):
        evento = {"tipo": tipo, "msg": msg}
        _jobs[job_id]["logs"].append(evento)
        _enviar(loop, queue, evento)

    try:
        if api_key.strip():
            os.environ["ANTHROPIC_API_KEY"] = api_key

        from analisador import (
            analisar_dissertacao,
            gerar_relatorio_analise,
            baixar_google_drive,
            gerar_dissertacao_revisada,
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

        resultado = analisar_dissertacao(
            diss_path=str(diss_final),
            api_key=api_key,
            log_fn=log,
        )

        log("📊 Gerando relatório…", "etapa")
        rel_dir = Path(tmpdir) / "relatorio"
        gerar_relatorio_analise(resultado, rel_dir)

        _jobs[job_id]["relatorio_dir"] = str(rel_dir)

        # ── Geração da dissertação revisada (.docx) ─────────────────────
        docx_path = None
        try:
            capitulos_raw = resultado.get("_capitulos_raw") or []
            docx_path = gerar_dissertacao_revisada(
                capitulos=capitulos_raw,
                resultados_caps=resultado.get("capitulos") or [],
                metadados=resultado.get("metadados") or {},
                api_key=api_key,
                output_path=rel_dir / "dissertacao_revisada.docx",
                log_fn=log,
            )
        except Exception as e:
            log(f"   ⚠ Erro ao gerar dissertação revisada: {e}", "warn")

        _jobs[job_id]["docx_revisado"] = str(docx_path) if docx_path else None
        _jobs[job_id]["status"] = "concluido"

        caps = resultado.get("capitulos") or []
        ponts = [c["analise"].get("pontuacao_geral", "") for c in caps]

        log_concluido = {
            "tipo": "concluido",
            "msg": "Análise concluída!",
            "resumo": {
                "n_capitulos":   len(caps),
                "aprovados":     sum(1 for p in ponts if p == "APROVADO"),
                "ressalvas":     sum(1 for p in ponts if p == "APROVADO_COM_RESSALVAS"),
                "revisao":       sum(1 for p in ponts if p == "REQUER_REVISAO"),
                "autor":         (resultado.get("metadados") or {}).get("autor") or "—",
                "titulo":        (resultado.get("metadados") or {}).get("titulo") or "—",
                "docx_disponivel": docx_path is not None,
            },
        }
        _jobs[job_id]["logs"].append(log_concluido)
        _enviar(loop, queue, log_concluido)

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
    loc_job_id: str,
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

        material_preimportado = None
        loc_job = _jobs.get(loc_job_id) if loc_job_id else None
        if loc_job and loc_job.get("status") == "concluido":
            relevantes = loc_job.get("arquivos_relevantes") or []
            material_texto = loc_job.get("material_texto") or {}
            material_preimportado = {n: material_texto[n] for n in relevantes if n in material_texto}
            log(f"📥 Usando {len(material_preimportado)} arquivo(s) importado(s) da Localização de Referências", "etapa")

        resultado = analisar(
            diss_path=diss_path,
            refs_dir=refs_dir,
            api_key=api_key,
            sem_verificacao=sem_verificacao,
            log_fn=log,
            material_preimportado=material_preimportado,
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
