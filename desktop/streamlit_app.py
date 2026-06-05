#!/usr/bin/env python3
"""
Verificador de Citações Acadêmicas — Interface Streamlit (v8.6)
"""
import os
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

# Garante que analisar.py seja encontrado mesmo em subpastas
_here = Path(__file__).parent
if str(_here) not in sys.path:
    sys.path.insert(0, str(_here))

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Verificador de Citações",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* ── Fundo e texto base ── */
[data-testid="stAppViewContainer"] { background: #eef0f5 !important; }
[data-testid="stAppViewContainer"] * { color: #1a1a1a; }
.main .block-container { max-width: 720px; padding-top: .5rem; }

/* ── Cabeçalho ── */
.cabecalho {
    background: linear-gradient(135deg, #1a2744 0%, #2c3e6b 100%);
    padding: 1.4em 1.6em 1.2em;
    border-radius: 14px;
    margin-bottom: 1.2em;
}
.cabecalho h1 { margin:0; font-size:1.5em; font-weight:800;
                letter-spacing:-.01em; color:#fff !important; }
.cabecalho p  { margin:.3em 0 0; font-size:.88em;
                color:rgba(255,255,255,.8) !important; }

/* ── Cards de passo ── */
.passo {
    background: #fff;
    border-radius: 12px;
    padding: 1.3em 1.5em;
    margin: .7em 0;
    box-shadow: 0 1px 6px rgba(0,0,0,.1);
    border-left: 5px solid #1a2744;
    color: #1a1a1a !important;
}
.passo.opcional { border-left-color: #7f8c8d; }
.passo h2, .passo h3, .passo p, .passo label,
.passo span, .passo small { color: #1a1a1a !important; }

/* ── Cards de resumo ── */
.resumo-box {
    background: #fff;
    border-radius: 10px;
    padding: 1em 1.2em;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
}
.resumo-num { font-size: 2.2em; font-weight: 800; line-height: 1; }
.resumo-txt { font-size: .78em; color: #555 !important; margin-top: .1em; }

/* ── Botão principal ── */
[data-testid="stBaseButton-primary"] button {
    background: #27ae60 !important;
    color: #fff !important;
    border: none !important;
    font-size: 1.1em !important;
    font-weight: 700 !important;
    height: 3.2em !important;
    border-radius: 10px !important;
}

/* ── Abas ── */
[data-baseweb="tab-list"] { gap: .3em; }
[data-baseweb="tab"] { font-size: .9em !important; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# CABEÇALHO
# ════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="cabecalho">
  <h1>📚 Verificador de Citações Acadêmicas</h1>
  <p>Lê a dissertação · Cruza referências · Verifica coerência com IA · Gera relatório</p>
</div>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PASSO 1 — DISSERTAÇÃO
# ════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="passo">', unsafe_allow_html=True)
st.subheader("1️⃣  Dissertação (.docx)")
docx_file = st.file_uploader(
    "Selecione o arquivo Word da dissertação",
    type=["docx", "doc"],
    key="docx",
    label_visibility="collapsed",
)
if docx_file:
    st.success(f"✅ **{docx_file.name}** — {docx_file.size // 1024} KB")
st.markdown("</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PASSO 2 — OBRAS DE REFERÊNCIA
# ════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="passo">', unsafe_allow_html=True)
st.subheader("2️⃣  Obras de referência")
st.caption("Escolha uma das opções abaixo:")

aba_arqs, aba_zip = st.tabs([
    "📄  Selecionar arquivos (PDFs/DOCXs)",
    "📦  Enviar como ZIP",
])

with aba_arqs:
    st.markdown("Selecione todos os PDFs e DOCXs da pasta de uma vez. "
                "No Android, abra o pen drive pelo seletor de arquivos.")
    arquivos_refs = st.file_uploader(
        "Selecione os arquivos de referência",
        type=["pdf", "docx", "doc", "txt"],
        accept_multiple_files=True,
        key="refs_arqs",
        label_visibility="collapsed",
    )
    if arquivos_refs:
        total_mb = sum(f.size for f in arquivos_refs) / (1024 * 1024)
        st.success(f"✅ **{len(arquivos_refs)} arquivo(s)** selecionado(s) — {total_mb:.1f} MB total")

with aba_zip:
    st.markdown("Compacte a pasta inteira em **.zip** e envie aqui (até 1 GB).")
    with st.expander("ℹ️  Como criar o ZIP no Android"):
        st.markdown("""
1. Instale o **ZArchiver** (gratuito, Play Store)
2. Navegue até a pasta com os PDFs no pen drive
3. Segure a pasta → **Comprimir** → formato **ZIP** → OK
4. Envie o `.zip` aqui
""")
    zip_file = st.file_uploader(
        "Arquivo .zip com as obras",
        type=["zip"],
        key="refs_zip",
        label_visibility="collapsed",
    )
    if zip_file:
        st.success(f"✅ **{zip_file.name}** — {zip_file.size / (1024 * 1024):.1f} MB")

st.markdown("</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PASSO 3 — CHAVE API (OPCIONAL)
# ════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="passo opcional">', unsafe_allow_html=True)
st.subheader("3️⃣  Chave API Claude *(opcional)*")
st.caption(
    "Necessária para verificar se os argumentos citados batem com o que as obras dizem. "
    "Sem ela, o sistema faz apenas o cruzamento de citações e referências."
)
api_key = st.text_input(
    "Chave API",
    type="password",
    placeholder="sk-ant-api03-…",
    label_visibility="collapsed",
)
sem_v = st.checkbox(
    "⚡  Apenas cruzamento — sem verificação por IA (mais rápido)",
    value=not bool(api_key),
)
st.markdown("</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# BOTÃO INICIAR
# ════════════════════════════════════════════════════════════════════════════

st.markdown("<br>", unsafe_allow_html=True)
rodar = st.button("▶  Iniciar análise", type="primary", use_container_width=True)

if rodar:
    arquivos_refs = st.session_state.get("refs_arqs") or []
    zip_file      = st.session_state.get("refs_zip")

    if not docx_file:
        st.error("❌  Selecione o arquivo da dissertação no passo 1.")
        st.stop()
    if not arquivos_refs and not zip_file:
        st.error("❌  Forneça as obras: selecione os arquivos ou envie um ZIP (passo 2).")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:

        # ── Salva DOCX ──────────────────────────────────────────────────────
        diss_path = os.path.join(tmpdir, "dissertacao.docx")
        with open(diss_path, "wb") as fh:
            fh.write(docx_file.getbuffer())

        refs_dir = os.path.join(tmpdir, "refs")
        os.makedirs(refs_dir, exist_ok=True)

        # ── Opção A: arquivos individuais ────────────────────────────────────
        if arquivos_refs:
            for arq in arquivos_refs:
                dest = os.path.join(refs_dir, arq.name)
                with open(dest, "wb") as fh:
                    fh.write(arq.getbuffer())

        # ── Opção B: extrai ZIP ──────────────────────────────────────────────
        elif zip_file:
            try:
                with zipfile.ZipFile(zip_file) as z:
                    for member in z.infolist():
                        nome = member.filename
                        if nome.startswith("__") or "/." in nome or nome.endswith("/"):
                            continue
                        z.extract(member, refs_dir)
            except zipfile.BadZipFile:
                st.error("❌  O arquivo enviado não é um ZIP válido.")
                st.stop()

        n_refs = sum(
            1 for p in Path(refs_dir).rglob("*")
            if p.suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
        )

        if n_refs == 0:
            st.error("❌  Nenhum arquivo de referência encontrado. Verifique o conteúdo enviado.")
            st.stop()

        # ── Importa analisar ─────────────────────────────────────────────────
        try:
            import analisar
        except ImportError:
            st.error("❌  Módulo analisar.py não encontrado. Contate o suporte.")
            st.stop()

        # ── Roda análise com log ao vivo ─────────────────────────────────────
        resultado = None
        with st.status(
            f"🔍  Analisando… ({n_refs} arquivo(s) de referência)",
            expanded=True,
        ) as status:
            def _log(msg):
                st.write(msg)

            resultado = analisar.analisar(
                diss_path=diss_path,
                refs_dir=refs_dir,
                api_key=api_key.strip(),
                log_fn=_log,
                sem_verificacao=sem_v,
            )

            if resultado:
                status.update(label="✅  Análise concluída!", state="complete", expanded=False)
            else:
                status.update(label="❌  Análise não produziu resultado", state="error")

        # ── Resultado ─────────────────────────────────────────────────────────
        if resultado:
            html = analisar.gerar_html(resultado)
            nome_rel = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.html"

            st.success("### ✅ Relatório pronto!")
            st.download_button(
                label="📥  Baixar relatório HTML",
                data=html,
                file_name=nome_rel,
                mime="text/html",
                use_container_width=True,
            )

            # ── Resumo visual ─────────────────────────────────────────────────
            st.markdown("---")
            st.subheader("📊 Resumo")

            cits  = resultado["citacoes"]
            refs  = resultado["referencias"]
            pares = resultado["pareamentos"]
            s_ref = resultado["citadas_sem_ref"]
            r_cit = resultado["refs_sem_citacao"]

            cols = st.columns(5)
            for col, (n, label, cor) in zip(cols, [
                (len(cits),  "Citações",      "#3498db"),
                (len(refs),  "Referências",   "#8e44ad"),
                (len(pares), "Pares OK",       "#27ae60"),
                (len(s_ref), "Cit. sem ref.", "#e74c3c"),
                (len(r_cit), "Ref. sem cit.", "#e67e22"),
            ]):
                col.markdown(
                    f'<div class="resumo-box" style="border-top:4px solid {cor}">'
                    f'<div class="resumo-num" style="color:{cor}">{n}</div>'
                    f'<div class="resumo-txt">{label}</div></div>',
                    unsafe_allow_html=True,
                )

            verifics = resultado.get("verificacoes", [])
            if verifics:
                corretas   = sum(1 for v in verifics if v.get("veredicto") == "CORRETO")
                incorretas = sum(1 for v in verifics if v.get("veredicto") == "INCORRETO")
                parciais   = sum(1 for v in verifics if v.get("veredicto") == "PARCIAL")
                sem_fonte  = sum(1 for v in verifics if v.get("veredicto") == "SEM_FONTE")

                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("🤖 Verificação semântica")
                cols2 = st.columns(4)
                for col, (n, label, cor) in zip(cols2, [
                    (corretas,   "✅ Corretas",   "#27ae60"),
                    (parciais,   "⚠️ Parciais",   "#f39c12"),
                    (incorretas, "❌ Incorretas", "#e74c3c"),
                    (sem_fonte,  "❓ Sem fonte",  "#95a5a6"),
                ]):
                    col.markdown(
                        f'<div class="resumo-box" style="border-top:4px solid {cor}">'
                        f'<div class="resumo-num" style="color:{cor}">{n}</div>'
                        f'<div class="resumo-txt">{label}</div></div>',
                        unsafe_allow_html=True,
                    )

            sugs = resultado.get("sugestoes_alternativas", [])
            com_match = [s for s in sugs if s.get("encontrou")]
            if com_match:
                st.info(f"💡 **{len(com_match)} obra(s) alternativa(s) sugerida(s)** "
                        f"— veja detalhes no relatório.")
