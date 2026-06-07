#!/usr/bin/env python3
"""
Verificador de Citações Acadêmicas — Interface Streamlit (v8.8)
"""
import io
import os
import re
import sys
import tempfile
import threading
import traceback
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

# ── Caminho local para importar analisar.py ──────────────────────────────────
_here = Path(__file__).parent
if str(_here) not in sys.path:
    sys.path.insert(0, str(_here))


def _run_job(job: dict, diss_path: str, refs_dir: str,
             api_key: str, sem_v: bool) -> None:
    """
    Roda em background thread.
    `job` é o mesmo dict guardado em st.session_state["job"].
    Qualquer modificação aqui fica visível no próximo fragment/rerun.
    """
    job["status"] = "rodando"

    def _log(msg: str) -> None:
        job["logs"].append(str(msg))

    try:
        import analisar as _an
        resultado = _an.analisar(
            diss_path=diss_path,
            refs_dir=refs_dir,
            api_key=api_key,
            log_fn=_log,
            sem_verificacao=sem_v,
        )
        job["resultado"] = resultado
        job["status"] = "concluido" if resultado else "vazio"
    except Exception:
        job["status"] = "erro"
        job["error"] = traceback.format_exc()


# ── Página ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Verificador de Citações",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.main .block-container { max-width: 740px; padding-top: 0; }
.hdr {
    background: linear-gradient(135deg,#1a2744,#2c4a8e);
    border-radius: 16px;
    padding: 1.5em 1.8em 1.3em;
    margin-bottom: 1.4em;
}
.hdr-titulo { font-size: 1.55em; font-weight: 800; color: #ffffff; margin: 0; letter-spacing: -0.02em; }
.hdr-sub    { font-size: .85em; color: rgba(255,255,255,.75); margin: .35em 0 0; }
.hdr-badge  { display: inline-block; background: rgba(255,255,255,.18); color: #fff;
              font-size: .72em; padding: .15em .65em; border-radius: 20px;
              margin-top: .6em; letter-spacing: .04em; }
.card-resultado { background: #fff; border-radius: 12px; padding: .9em 1.1em;
                  text-align: center; box-shadow: 0 1px 5px rgba(0,0,0,.09); }
.card-n   { font-size: 2.4em; font-weight: 800; line-height: 1; }
.card-txt { font-size: .75em; color: #555; margin-top: .2em; }
.rodape   { text-align: center; font-size: .75em; color: #999; margin-top: 2em;
            padding-top: 1em; border-top: 1px solid #ddd; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# CABEÇALHO
# ════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hdr">
  <div class="hdr-titulo">📚 Verificador de Citações Acadêmicas</div>
  <div class="hdr-sub">
    Lê a dissertação &nbsp;·&nbsp; Cruza referências &nbsp;·&nbsp;
    Verifica coerência com IA &nbsp;·&nbsp; Gera relatório
  </div>
  <div class="hdr-badge">v8.8</div>
</div>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# GESTÃO DE ARQUIVOS DE REFERÊNCIA — disco, não RAM
# ════════════════════════════════════════════════════════════════════════════

def _refs_dir() -> str:
    if "refs_tmpdir" not in st.session_state:
        st.session_state["refs_tmpdir"] = tempfile.mkdtemp(
            prefix=f"cit_{uuid.uuid4().hex[:8]}_"
        )
        st.session_state["refs_meta"] = {}
    return st.session_state["refs_tmpdir"]


def _add_file(nome: str, data: bytes) -> None:
    d = _refs_dir()
    with open(os.path.join(d, nome), "wb") as f:
        f.write(data)
    st.session_state["refs_meta"][nome] = len(data)


def _del_file(nome: str) -> None:
    d = _refs_dir()
    try:
        os.remove(os.path.join(d, nome))
    except FileNotFoundError:
        pass
    st.session_state["refs_meta"].pop(nome, None)


def _clear_files() -> None:
    meta = st.session_state.get("refs_meta", {})
    d    = st.session_state.get("refs_tmpdir", "")
    for nome in list(meta):
        try:
            os.remove(os.path.join(d, nome))
        except FileNotFoundError:
            pass
    st.session_state["refs_meta"] = {}


# Inicializa estruturas na primeira execução da sessão
_refs_dir()
if "resultado"     not in st.session_state:
    st.session_state["resultado"]     = None
if "job"           not in st.session_state:
    st.session_state["job"]           = None
if "analise_erro"  not in st.session_state:
    st.session_state["analise_erro"]  = None
if "analise_aviso" not in st.session_state:
    st.session_state["analise_aviso"] = None


# ════════════════════════════════════════════════════════════════════════════
# PAINEL DE PROGRESSO  (st.fragment — atualiza a cada 2 s sem sleep/bloquear)
# Isso resolve o "Connection error" em tablets: nenhum sleep no thread
# principal, o WebSocket nunca fica ocioso durante a análise.
# ════════════════════════════════════════════════════════════════════════════
@st.fragment(run_every=2)
def _painel_progresso() -> None:
    job = st.session_state.get("job")
    if job is None:
        return

    _status = job["status"]
    _logs   = job.get("logs", [])

    if _status in ("iniciando", "rodando"):
        _etapa = next(
            (l for l in reversed(_logs) if l.strip().startswith("ETAPA")),
            "Aguardando início…",
        )
        st.info(
            f"⏳ **{_etapa}**  \n"
            "A análise roda em segundo plano — pode minimizar o app. "
            "Painel atualiza a cada 2 s.",
            icon="🔄",
        )
        st.markdown(f"**📋 Log de progresso ({len(_logs)} linha(s)):**")
        if _logs:
            st.code("\n".join(_logs[-60:]), language="")
        else:
            st.caption("⏳ Iniciando análise… aguarde as primeiras mensagens.")
        return  # fragment re-executa sozinho — sem sleep, sem rerun manual

    # Análise concluída: guarda resultado e aciona rerun completo
    if _status == "concluido":
        st.session_state["resultado"] = job["resultado"]
        st.session_state["job"] = None
        st.rerun()

    elif _status == "vazio":
        st.session_state["analise_aviso"] = (
            "A análise foi executada mas não retornou dados.  \n"
            "Verifique se o DOCX contém citações no formato ABNT."
        )
        st.session_state["job"] = None
        st.rerun()

    elif _status == "erro":
        # Persiste o erro ANTES de limpar o job, para aparecer na página completa
        st.session_state["analise_erro"] = job.get("error", "Erro desconhecido")
        st.session_state["job"] = None
        st.rerun()


_painel_progresso()

# Enquanto análise roda, esconde o formulário abaixo
if st.session_state.get("job") is not None:
    st.stop()


# ════════════════════════════════════════════════════════════════════════════
# MENSAGENS PERSISTENTES (erro ou aviso da última análise)
# ════════════════════════════════════════════════════════════════════════════
if st.session_state.get("analise_erro"):
    st.error("❌  **Erro durante a análise:**")
    st.code(st.session_state["analise_erro"], language="")
    if st.button("🗑️  Fechar erro", key="btn_fechar_erro"):
        st.session_state["analise_erro"] = None
        st.rerun()

if st.session_state.get("analise_aviso"):
    st.warning(f"⚠️ {st.session_state['analise_aviso']}")
    if st.button("🗑️  Fechar aviso", key="btn_fechar_aviso"):
        st.session_state["analise_aviso"] = None
        st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# PASSO 1 — DISSERTAÇÃO
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 1️⃣ &nbsp;Dissertação")
    docx_file = st.file_uploader(
        "Selecione o arquivo **.docx** da dissertação",
        type=["docx", "doc"],
        key="docx",
    )
    if docx_file:
        st.success(f"✅ **{docx_file.name}** &nbsp;·&nbsp; {docx_file.size // 1024} KB")


# ════════════════════════════════════════════════════════════════════════════
# PASSO 2 — OBRAS DE REFERÊNCIA
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 2️⃣ &nbsp;Obras de referência")

    aba_drive, aba_zip, aba_multi, aba_um = st.tabs([
        "☁️ Google Drive (700 MB+)",
        "📦 ZIP local",
        "🗂️ Vários arquivos",
        "➕ Um por vez",
    ])

    # ── ABA GOOGLE DRIVE ─────────────────────────────────────────────────────
    with aba_drive:
        st.caption(
            "**Melhor opção para arquivos grandes (ex: 700 MB).**  \n"
            "Suba o ZIP para o Google Drive, compartilhe com "
            "\"Qualquer pessoa com o link\" e cole o link abaixo. "
            "O servidor baixa direto — sem passar pelo tablet."
        )
        st.markdown(
            "**Como obter o link:**  \n"
            "1. Google Drive → botão direito no ZIP &nbsp;·&nbsp; "
            "2. **Compartilhar → Qualquer pessoa com o link** &nbsp;·&nbsp; "
            "3. Copiar link e colar aqui"
        )
        drive_url = st.text_input(
            "Link do Google Drive",
            placeholder="https://drive.google.com/file/d/…/view",
            label_visibility="collapsed",
        )
        if st.button("⬇️  Baixar do Drive e extrair", key="btn_drive", use_container_width=True):
            if not drive_url.strip():
                st.warning("Cole o link do Google Drive primeiro.")
            else:
                try:
                    import gdown
                    _url = drive_url.strip()
                    _m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', _url)
                    if not _m:
                        _m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', _url)
                    _dl_url = (
                        f"https://drive.google.com/uc?export=download&id={_m.group(1)}"
                        if _m else _url
                    )
                    with st.status("⬇️  Baixando do Google Drive…", expanded=True) as _st:
                        st.write("Conectando ao Drive…")
                        with tempfile.TemporaryDirectory() as _td:
                            _out = os.path.join(_td, "refs_drive.zip")
                            gdown.download(_dl_url, _out, quiet=True)
                            if not os.path.exists(_out) or os.path.getsize(_out) < 100:
                                _st.update(label="❌  Falha no download", state="error")
                                st.error(
                                    "Não foi possível baixar. Verifique:\n"
                                    "- Link público (\"Qualquer pessoa com o link\")\n"
                                    "- Arquivo é um ZIP válido"
                                )
                            else:
                                st.write("Extraindo arquivos para o disco…")
                                with zipfile.ZipFile(_out) as _z:
                                    validos = [
                                        n for n in _z.namelist()
                                        if Path(n).suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
                                        and not Path(n).name.startswith((".", "__"))
                                    ]
                                    for nz in validos:
                                        _add_file(Path(nz).name, _z.read(nz))
                                _st.update(
                                    label=f"✅  {len(validos)} arquivo(s) prontos!",
                                    state="complete", expanded=False,
                                )
                                st.rerun()
                except Exception as e:
                    st.error(f"❌  Erro: {e}")

    # ── ABA ZIP LOCAL ─────────────────────────────────────────────────────────
    with aba_zip:
        st.caption("Compacte a pasta em ZIP no computador e envie. Ideal para muitos arquivos.")
        zip_file = st.file_uploader(
            "Selecione o arquivo **.zip**",
            type=["zip"],
            key="refs_zip",
            label_visibility="collapsed",
        )
        if zip_file is not None:
            try:
                with zipfile.ZipFile(io.BytesIO(zip_file.read())) as z:
                    validos = [
                        n for n in z.namelist()
                        if Path(n).suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
                        and not Path(n).name.startswith((".", "__"))
                    ]
                    if validos:
                        for nz in validos:
                            _add_file(Path(nz).name, z.read(nz))
                        st.success(f"✅ **{len(validos)} arquivo(s)** extraídos e prontos!")
                    else:
                        st.warning("⚠️ Nenhum PDF/DOCX/TXT encontrado no ZIP.")
            except zipfile.BadZipFile:
                st.error("❌ ZIP inválido ou corrompido.")
            except Exception as e:
                st.error(f"❌ Erro: {e}")

    # ── ABA VÁRIOS ARQUIVOS ───────────────────────────────────────────────────
    with aba_multi:
        st.caption("Selecione vários arquivos de uma vez (Ctrl/Shift no computador).")
        multi_arqs = st.file_uploader(
            "Selecione PDFs e DOCXs",
            type=["pdf", "docx", "doc", "txt"],
            accept_multiple_files=True,
            key="refs_multi",
            label_visibility="collapsed",
        )
        if multi_arqs:
            novos = [a for a in multi_arqs if a.name not in st.session_state["refs_meta"]]
            for arq in novos:
                _add_file(arq.name, arq.getbuffer().tobytes())
            if novos:
                st.success(f"✅ **{len(novos)} arquivo(s) novos** adicionados!")

    # ── ABA UM POR VEZ ────────────────────────────────────────────────────────
    with aba_um:
        st.caption("Envie um arquivo por vez — útil quando os outros métodos não funcionam.")
        novo_arq = st.file_uploader(
            "Selecione um arquivo",
            type=["pdf", "docx", "doc", "txt"],
            key="refs_individual",
            label_visibility="collapsed",
        )
        if novo_arq is not None:
            _add_file(novo_arq.name, novo_arq.getbuffer().tobytes())

    # ── LISTA DE ARQUIVOS ─────────────────────────────────────────────────────
    meta     = st.session_state["refs_meta"]
    n_total  = len(meta)
    mb_total = sum(meta.values()) / (1024 * 1024)

    st.markdown("---")
    if n_total > 0:
        st.success(f"✅ **{n_total} arquivo(s)** prontos para análise — {mb_total:.1f} MB total")
        with st.expander(f"Ver lista ({n_total} arquivos)"):
            for nome, tam in list(meta.items()):
                col_nome, col_del = st.columns([5, 1])
                col_nome.markdown(
                    f"📄 {nome} &nbsp;<small>({tam // 1024} KB)</small>",
                    unsafe_allow_html=True,
                )
                if col_del.button("✕", key=f"del_{nome}", help="Remover"):
                    _del_file(nome)
                    st.rerun()
        if st.button("🗑️  Limpar todos os arquivos", use_container_width=True):
            _clear_files()
            st.rerun()
    else:
        st.info("Nenhum arquivo adicionado ainda. Use uma das abas acima.", icon="📂")


# ════════════════════════════════════════════════════════════════════════════
# PASSO 3 — CHAVE API (OPCIONAL)
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown(
        "### 3️⃣ &nbsp;Chave API Claude "
        "&nbsp;<small style='font-weight:400;color:#888'>opcional</small>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Necessária para verificação semântica — verifica se os argumentos "
        "citados batem com o que as obras dizem. "
        "Sem ela o sistema faz só o cruzamento."
    )
    api_key = st.text_input(
        "Chave API",
        type="password",
        placeholder="sk-ant-api03-…",
        label_visibility="collapsed",
    )
    sem_v = st.checkbox(
        "⚡  Apenas cruzamento — sem verificação por IA (mais rápido)",
        value=(not bool(api_key)),
    )


# ════════════════════════════════════════════════════════════════════════════
# BOTÃO INICIAR
# ════════════════════════════════════════════════════════════════════════════
st.markdown("<br>", unsafe_allow_html=True)

rodar = st.button(
    "▶  Iniciar análise",
    type="primary",
    use_container_width=True,
)

if rodar:
    meta = st.session_state.get("refs_meta", {})

    if not docx_file:
        st.error("❌  Selecione o arquivo da dissertação (passo 1).")
        st.stop()
    if not meta:
        st.error("❌  Adicione pelo menos um arquivo de referência (passo 2).")
        st.stop()

    try:
        import analisar  # noqa: F401
    except ImportError:
        st.error("❌  analisar.py não encontrado. Contate o suporte.")
        st.stop()

    refs_dir = st.session_state["refs_tmpdir"]
    n_refs = sum(
        1 for p in Path(refs_dir).rglob("*")
        if p.suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
    )
    if n_refs == 0:
        st.error("❌  Nenhum arquivo de referência encontrado.")
        st.stop()

    diss_path = os.path.join(refs_dir, "_dissertacao.docx")
    with open(diss_path, "wb") as fh:
        fh.write(docx_file.getbuffer())

    # Limpa mensagens anteriores
    st.session_state["analise_erro"]  = None
    st.session_state["analise_aviso"] = None

    # Cria job e guarda no session_state.
    # A thread recebe referência direta ao mesmo dict.
    job = {"status": "iniciando", "logs": [], "resultado": None, "error": None}
    st.session_state["job"] = job

    threading.Thread(
        target=_run_job,
        args=(job, diss_path, refs_dir, api_key.strip(), sem_v),
        daemon=True,
    ).start()

    st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# EXIBE RESULTADO
# ════════════════════════════════════════════════════════════════════════════
resultado = st.session_state.get("resultado")
if resultado:
    try:
        import analisar
        html = analisar.gerar_html(resultado)
    except Exception as _exc:
        st.error(f"❌  Erro ao gerar relatório HTML: {_exc}")
        html = None

    if html:
        nome_rel = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.html"

        st.success("### ✅  Relatório pronto!")
        st.download_button(
            label="📥  Baixar relatório HTML",
            data=html,
            file_name=nome_rel,
            mime="text/html",
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("### 📊 Resumo da análise")
        resumo = [
            (len(resultado["citacoes"]),        "Citações encontradas", "#3498db"),
            (len(resultado["referencias"]),     "Referências na lista", "#8e44ad"),
            (len(resultado["pareamentos"]),     "Pares citação↔ref",    "#27ae60"),
            (len(resultado["citadas_sem_ref"]), "Citadas sem referência","#e74c3c"),
            (len(resultado["refs_sem_citacao"]),"Refs sem citação",     "#e67e22"),
        ]
        cols = st.columns(len(resumo))
        for col, (n, label, cor) in zip(cols, resumo):
            col.markdown(
                f'<div class="card-resultado" style="border-top:4px solid {cor}">'
                f'<div class="card-n" style="color:{cor}">{n}</div>'
                f'<div class="card-txt">{label}</div></div>',
                unsafe_allow_html=True,
            )

        verifics = resultado.get("verificacoes", [])
        if verifics:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 🤖 Verificação semântica")
            corretas   = sum(1 for v in verifics if v.get("veredicto") == "CORRETO")
            incorretas = sum(1 for v in verifics if v.get("veredicto") == "INCORRETO")
            parciais   = sum(1 for v in verifics if v.get("veredicto") == "PARCIAL")
            sem_fonte  = sum(1 for v in verifics if v.get("veredicto") == "SEM_FONTE")
            cols2 = st.columns(4)
            for col, (n, label, cor) in zip(cols2, [
                (corretas,   "Corretas",   "#27ae60"),
                (parciais,   "Parciais",   "#f39c12"),
                (incorretas, "Incorretas", "#e74c3c"),
                (sem_fonte,  "Sem fonte",  "#95a5a6"),
            ]):
                col.markdown(
                    f'<div class="card-resultado" style="border-top:4px solid {cor}">'
                    f'<div class="card-n" style="color:{cor}">{n}</div>'
                    f'<div class="card-txt">{label}</div></div>',
                    unsafe_allow_html=True,
                )

        sugs = [s for s in resultado.get("sugestoes_alternativas", []) if s.get("encontrou")]
        if sugs:
            st.info(
                f"💡 **{len(sugs)} obra(s) alternativa(s) sugerida(s)** "
                f"para citações sem referência — detalhes no relatório.",
                icon="💡",
            )

st.markdown(
    '<div class="rodape">Verificador de Citações Acadêmicas v8.8 · '
    'Desenvolvido com Claude API</div>',
    unsafe_allow_html=True,
)
