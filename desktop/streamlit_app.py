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
/* Fundo geral */
[data-testid="stAppViewContainer"] { background: #f0f2f6; }

/* Cabeçalho do app */
.cabecalho {
    background: linear-gradient(135deg, #1a2744 0%, #2c3e6b 100%);
    color: white;
    padding: 1.4em 1.6em 1.2em;
    border-radius: 14px;
    margin-bottom: 1.2em;
}
.cabecalho h1 { margin: 0; font-size: 1.5em; font-weight: 800; letter-spacing: -.01em; }
.cabecalho p  { margin: .3em 0 0; font-size: .88em; opacity: .8; }

/* Cards de passo */
.passo {
    background: white;
    border-radius: 12px;
    padding: 1.3em 1.5em;
    margin: .7em 0;
    box-shadow: 0 1px 6px rgba(0,0,0,.08);
    border-left: 5px solid #1a2744;
}
.passo.opcional { border-left-color: #aab; }

/* Métricas de resumo */
.resumo-box {
    background: white;
    border-radius: 10px;
    padding: 1em 1.2em;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
}
.resumo-num  { font-size: 2.2em; font-weight: 800; line-height: 1; }
.resumo-txt  { font-size: .78em; color: #666; margin-top: .1em; }

/* Botão principal */
[data-testid="stBaseButton-primary"] > button {
    background: #27ae60 !important;
    border: none !important;
    font-size: 1.1em !important;
    font-weight: 700 !important;
    height: 3.2em !important;
    border-radius: 10px !important;
    letter-spacing: .02em !important;
}

/* Ajuste largura */
.main .block-container { max-width: 720px; padding-top: .5rem; }
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

aba_drive, aba_zip = st.tabs(["☁️  Google Drive (recomendado — sem limite de tamanho)", "📦  Upload ZIP (até 1 GB)"])

with aba_drive:
    st.markdown(
        "Compartilhe a pasta do Google Drive e cole o link aqui. "
        "O sistema vai baixar os arquivos automaticamente."
    )
    with st.expander("ℹ️  Como compartilhar a pasta no Google Drive"):
        st.markdown("""
1. Copie a pasta de obras do **pen drive para o Google Drive** (pelo app Google Drive ou pelo app Arquivos do Android)
2. Abra o Google Drive no navegador → clique com botão direito (ou segure) na pasta → **Compartilhar**
3. Mude para **"Qualquer pessoa com o link"** → copie o link
4. Cole o link abaixo
""")
    drive_url = st.text_input(
        "Link do Google Drive",
        placeholder="https://drive.google.com/drive/folders/...",
        label_visibility="collapsed",
        key="drive_url",
    )
    if drive_url:
        st.success("✅ Link recebido — os arquivos serão baixados ao iniciar a análise.")

with aba_zip:
    st.markdown("Para pastas **menores que 1 GB**. Compacte a pasta em ZIP antes de enviar.")
    with st.expander("ℹ️  Como criar o ZIP no tablet Android"):
        st.markdown("""
1. Instale o **ZArchiver** (gratuito, Play Store)
2. Navegue até a pasta com os PDFs
3. Segure a pasta → **Comprimir** → formato **ZIP** → OK
4. Envie o arquivo `.zip` aqui
""")
    zip_file = st.file_uploader(
        "Arquivo .zip com as obras",
        type=["zip"],
        key="refs_zip",
        label_visibility="collapsed",
    )
    if zip_file:
        st.success(f"✅ **{zip_file.name}** — {zip_file.size / (1024*1024):.1f} MB")

st.markdown("</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PASSO 3 — CHAVE API (OPCIONAL)
# ════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="passo opcional">', unsafe_allow_html=True)
st.subheader("3️⃣  Chave API Claude *(opcional)*")
st.caption(
    "Necessária para verificar se os argumentos citados batem com o que as obras dizem. "
    "Sem ela, o sistema só faz o cruzamento de citações e referências."
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
    drive_url = st.session_state.get("drive_url", "").strip()
    zip_file  = st.session_state.get("refs_zip")

    if not docx_file:
        st.error("❌  Selecione o arquivo da dissertação no passo 1.")
        st.stop()
    if not drive_url and not zip_file:
        st.error("❌  Forneça as obras de referência: link do Google Drive ou arquivo ZIP (passo 2).")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:

        # ── Salva DOCX ──────────────────────────────────────────────────────
        diss_path = os.path.join(tmpdir, "dissertacao.docx")
        with open(diss_path, "wb") as fh:
            fh.write(docx_file.getbuffer())

        refs_dir = os.path.join(tmpdir, "refs")
        os.makedirs(refs_dir, exist_ok=True)

        # ── Opção A: baixa do Google Drive ──────────────────────────────────
        if drive_url:
            try:
                import subprocess
                with st.spinner("☁️  Baixando arquivos do Google Drive…"):
                    r = subprocess.run(
                        ["gdown", "--folder", drive_url, "-O", refs_dir],
                        capture_output=True, text=True, timeout=600,
                    )
                if r.returncode != 0:
                    st.error(f"❌  Erro ao baixar do Drive:\n```\n{r.stderr[-400:]}\n```")
                    st.stop()
            except FileNotFoundError:
                st.error("❌  'gdown' não encontrado. Verifique o requirements.txt.")
                st.stop()

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

        # ── Importa analisar ─────────────────────────────────────────────
        try:
            import analisar
        except ImportError:
            st.error("❌  Módulo analisar.py não encontrado. Contate o suporte.")
            st.stop()

        # ── Roda análise com log ao vivo ─────────────────────────────────
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

        # ── Resultado ────────────────────────────────────────────────────
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

            # ── Resumo visual ────────────────────────────────────────────
            st.markdown("---")
            st.subheader("📊 Resumo")

            cits   = resultado["citacoes"]
            refs   = resultado["referencias"]
            pares  = resultado["pareamentos"]
            s_ref  = resultado["citadas_sem_ref"]
            r_cit  = resultado["refs_sem_citacao"]
            verifics = resultado.get("verificacoes", [])

            cols = st.columns(5)
            dados = [
                (len(cits),  "Citações",       "#3498db"),
                (len(refs),  "Referências",    "#8e44ad"),
                (len(pares), "Pares OK",        "#27ae60"),
                (len(s_ref), "Cit. sem ref.",   "#e74c3c"),
                (len(r_cit), "Ref. sem cit.",   "#e67e22"),
            ]
            for col, (n, label, cor) in zip(cols, dados):
                col.markdown(
                    f'<div class="resumo-box" style="border-top:4px solid {cor}">'
                    f'<div class="resumo-num" style="color:{cor}">{n}</div>'
                    f'<div class="resumo-txt">{label}</div></div>',
                    unsafe_allow_html=True,
                )

            if verifics:
                corretas   = sum(1 for v in verifics if v.get("veredicto") == "CORRETO")
                incorretas = sum(1 for v in verifics if v.get("veredicto") == "INCORRETO")
                parciais   = sum(1 for v in verifics if v.get("veredicto") == "PARCIAL")
                sem_fonte  = sum(1 for v in verifics if v.get("veredicto") == "SEM_FONTE")

                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("🤖 Verificação semântica")
                cols2 = st.columns(4)
                sem_dados = [
                    (corretas,   "✅ Corretas",   "#27ae60"),
                    (parciais,   "⚠️ Parciais",   "#f39c12"),
                    (incorretas, "❌ Incorretas", "#e74c3c"),
                    (sem_fonte,  "❓ Sem fonte",  "#95a5a6"),
                ]
                for col, (n, label, cor) in zip(cols2, sem_dados):
                    col.markdown(
                        f'<div class="resumo-box" style="border-top:4px solid {cor}">'
                        f'<div class="resumo-num" style="color:{cor}">{n}</div>'
                        f'<div class="resumo-txt">{label}</div></div>',
                        unsafe_allow_html=True,
                    )

            sugs = resultado.get("sugestoes_alternativas", [])
            com_match = [s for s in sugs if s.get("encontrou")]
            if com_match:
                st.info(f"💡 **{len(com_match)} obra(s) alternativa(s) sugerida(s)** — veja detalhes no relatório.")
