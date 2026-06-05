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

_here = Path(__file__).parent
if str(_here) not in sys.path:
    sys.path.insert(0, str(_here))

# ── Página ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Verificador de Citações",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# CSS mínimo: só afeta elementos HTML que eu mesmo crio
st.markdown("""
<style>
.main .block-container { max-width: 740px; padding-top: 0; }

.hdr {
    background: linear-gradient(135deg,#1a2744,#2c4a8e);
    border-radius: 16px;
    padding: 1.5em 1.8em 1.3em;
    margin-bottom: 1.4em;
}
.hdr-titulo {
    font-size: 1.55em;
    font-weight: 800;
    color: #ffffff;
    margin: 0;
    letter-spacing: -0.02em;
}
.hdr-sub {
    font-size: .85em;
    color: rgba(255,255,255,.75);
    margin: .35em 0 0;
}
.hdr-badge {
    display: inline-block;
    background: rgba(255,255,255,.18);
    color: #fff;
    font-size: .72em;
    padding: .15em .65em;
    border-radius: 20px;
    margin-top: .6em;
    letter-spacing: .04em;
}

.card-resultado {
    background: #fff;
    border-radius: 12px;
    padding: .9em 1.1em;
    text-align: center;
    box-shadow: 0 1px 5px rgba(0,0,0,.09);
}
.card-n   { font-size: 2.4em; font-weight: 800; line-height: 1; }
.card-txt { font-size: .75em; color: #555; margin-top: .2em; }

.rodape {
    text-align: center;
    font-size: .75em;
    color: #999;
    margin-top: 2em;
    padding-top: 1em;
    border-top: 1px solid #ddd;
}
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
  <div class="hdr-badge">v8.6</div>
</div>
""", unsafe_allow_html=True)


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
    st.caption("Escolha como enviar os PDFs/DOCXs das obras citadas:")

    aba_arqs, aba_zip = st.tabs([
        "📄  Selecionar arquivos",
        "📦  Enviar ZIP",
    ])

    with aba_arqs:
        st.info(
            "Selecione todos os PDFs e DOCXs de uma vez.\n\n"
            "**No tablet:** abra o seletor de arquivos → navegue até o pen drive → "
            "segure um arquivo para entrar no modo de seleção múltipla → selecione todos.",
            icon="💡",
        )
        arquivos_refs = st.file_uploader(
            "PDFs e DOCXs das obras",
            type=["pdf", "docx", "doc", "txt"],
            accept_multiple_files=True,
            key="refs_arqs",
            label_visibility="collapsed",
        )
        if arquivos_refs:
            total_mb = sum(f.size for f in arquivos_refs) / (1024 * 1024)
            st.success(
                f"✅ **{len(arquivos_refs)} arquivo(s)** selecionado(s) "
                f"— {total_mb:.1f} MB no total"
            )

    with aba_zip:
        st.info(
            "Compacte a pasta inteira em **.zip** e envie aqui (limite: 1 GB).",
            icon="📦",
        )
        with st.expander("Como criar o ZIP no Android (ZArchiver)"):
            st.markdown("""
1. Instale o **ZArchiver** (gratuito — Play Store)
2. Abra o ZArchiver, navegue até a pasta dos PDFs no pen drive
3. Segure a pasta → toque em **Comprimir** → formato **ZIP** → OK
4. Selecione o arquivo `.zip` gerado aqui abaixo
""")
        zip_file = st.file_uploader(
            "Arquivo .zip com as obras",
            type=["zip"],
            key="refs_zip",
            label_visibility="collapsed",
        )
        if zip_file:
            st.success(
                f"✅ **{zip_file.name}** — {zip_file.size / (1024 * 1024):.1f} MB"
            )


# ════════════════════════════════════════════════════════════════════════════
# PASSO 3 — CHAVE API (OPCIONAL)
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 3️⃣ &nbsp;Chave API Claude &nbsp;<small style='font-weight:400;color:#888'>opcional</small>", unsafe_allow_html=True)
    st.caption(
        "Necessária para a verificação semântica — verifica se os argumentos "
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
    arquivos_refs = st.session_state.get("refs_arqs") or []
    zip_file      = st.session_state.get("refs_zip")

    if not docx_file:
        st.error("❌  Selecione o arquivo da dissertação (passo 1).")
        st.stop()
    if not arquivos_refs and not zip_file:
        st.error("❌  Forneça as obras de referência — arquivos ou ZIP (passo 2).")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:

        # Salva DOCX
        diss_path = os.path.join(tmpdir, "dissertacao.docx")
        with open(diss_path, "wb") as fh:
            fh.write(docx_file.getbuffer())

        refs_dir = os.path.join(tmpdir, "refs")
        os.makedirs(refs_dir, exist_ok=True)

        # Opção A: arquivos individuais
        if arquivos_refs:
            for arq in arquivos_refs:
                with open(os.path.join(refs_dir, arq.name), "wb") as fh:
                    fh.write(arq.getbuffer())

        # Opção B: ZIP
        elif zip_file:
            try:
                with zipfile.ZipFile(zip_file) as z:
                    for member in z.infolist():
                        nome = member.filename
                        if nome.startswith("__") or "/." in nome or nome.endswith("/"):
                            continue
                        z.extract(member, refs_dir)
            except zipfile.BadZipFile:
                st.error("❌  O arquivo não é um ZIP válido.")
                st.stop()

        n_refs = sum(
            1 for p in Path(refs_dir).rglob("*")
            if p.suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
        )
        if n_refs == 0:
            st.error("❌  Nenhum arquivo de referência encontrado. Verifique o conteúdo.")
            st.stop()

        try:
            import analisar
        except ImportError:
            st.error("❌  analisar.py não encontrado. Contate o suporte.")
            st.stop()

        resultado = None
        with st.status(
            f"🔍  Analisando… ({n_refs} arquivo(s) de referência)",
            expanded=True,
        ) as status:
            resultado = analisar.analisar(
                diss_path=diss_path,
                refs_dir=refs_dir,
                api_key=api_key.strip(),
                log_fn=st.write,
                sem_verificacao=sem_v,
            )
            if resultado:
                status.update(
                    label="✅  Análise concluída!",
                    state="complete",
                    expanded=False,
                )
            else:
                status.update(label="❌  Análise não produziu resultado", state="error")

        if resultado:
            html     = analisar.gerar_html(resultado)
            nome_rel = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.html"

            st.balloons()
            st.success("### ✅  Relatório pronto!")
            st.download_button(
                label="📥  Baixar relatório HTML",
                data=html,
                file_name=nome_rel,
                mime="text/html",
                use_container_width=True,
            )

            # ── Resumo ────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### 📊 Resumo da análise")

            resumo = [
                (len(resultado["citacoes"]),       "Citações encontradas", "#3498db"),
                (len(resultado["referencias"]),    "Referências na lista", "#8e44ad"),
                (len(resultado["pareamentos"]),    "Pares citação↔ref",    "#27ae60"),
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

                sem_dados = [
                    (corretas,   "Corretas",   "#27ae60"),
                    (parciais,   "Parciais",   "#f39c12"),
                    (incorretas, "Incorretas", "#e74c3c"),
                    (sem_fonte,  "Sem fonte",  "#95a5a6"),
                ]
                cols2 = st.columns(4)
                for col, (n, label, cor) in zip(cols2, sem_dados):
                    col.markdown(
                        f'<div class="card-resultado" style="border-top:4px solid {cor}">'
                        f'<div class="card-n" style="color:{cor}">{n}</div>'
                        f'<div class="card-txt">{label}</div></div>',
                        unsafe_allow_html=True,
                    )

            sugs = [s for s in resultado.get("sugestoes_alternativas", [])
                    if s.get("encontrou")]
            if sugs:
                st.info(
                    f"💡 **{len(sugs)} obra(s) alternativa(s) sugerida(s)** "
                    f"para citações sem referência — detalhes no relatório.",
                    icon="💡",
                )

st.markdown(
    '<div class="rodape">Verificador de Citações Acadêmicas v8.6 · '
    'Desenvolvido com Claude API</div>',
    unsafe_allow_html=True,
)
