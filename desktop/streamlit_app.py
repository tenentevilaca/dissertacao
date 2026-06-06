#!/usr/bin/env python3
"""
Verificador de Citações Acadêmicas — Interface Streamlit (v8.6)
"""
import io
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

if "refs_acumulados" not in st.session_state:
    st.session_state["refs_acumulados"] = {}  # nome → bytes

with st.container(border=True):
    st.markdown("### 2️⃣ &nbsp;Obras de referência")

    aba_zip, aba_multi, aba_um = st.tabs([
        "📦 ZIP (recomendado)",
        "🗂️ Vários arquivos",
        "➕ Um por vez",
    ])

    # ── ABA ZIP ──────────────────────────────────────────────────────────────
    with aba_zip:
        st.caption(
            "Compacte todos os PDFs/DOCXs numa pasta ZIP no computador "
            "e envie de uma vez. Ideal para muitos arquivos."
        )
        zip_file = st.file_uploader(
            "Selecione o arquivo **.zip**",
            type=["zip"],
            key="refs_zip",
            label_visibility="collapsed",
        )
        if zip_file is not None:
            try:
                zip_bytes = io.BytesIO(zip_file.read())
                with zipfile.ZipFile(zip_bytes) as z:
                    nomes_validos = [
                        n for n in z.namelist()
                        if Path(n).suffix.lower() in {".pdf", ".docx", ".doc", ".txt"}
                        and not Path(n).name.startswith(".")
                        and not Path(n).name.startswith("__")
                    ]
                    if nomes_validos:
                        adicionados = 0
                        for nome_zip in nomes_validos:
                            nome_base = Path(nome_zip).name
                            st.session_state["refs_acumulados"][nome_base] = z.read(nome_zip)
                            adicionados += 1
                        st.success(
                            f"✅ **{adicionados} arquivo(s)** extraídos do ZIP "
                            f"e adicionados à fila!"
                        )
                    else:
                        st.warning(
                            "⚠️ Nenhum PDF/DOCX/TXT encontrado dentro do ZIP. "
                            "Verifique o conteúdo do arquivo."
                        )
            except zipfile.BadZipFile:
                st.error("❌ Arquivo ZIP inválido ou corrompido. Tente compactar novamente.")
            except Exception as e:
                st.error(f"❌ Erro ao processar ZIP: {e}")

    # ── ABA VÁRIOS ARQUIVOS ───────────────────────────────────────────────────
    with aba_multi:
        st.caption(
            "Selecione múltiplos arquivos de uma vez "
            "(segure Ctrl ou Shift no computador, ou selecione um a um no Android)."
        )
        multi_arqs = st.file_uploader(
            "Selecione PDFs e DOCXs",
            type=["pdf", "docx", "doc", "txt"],
            accept_multiple_files=True,
            key="refs_multi",
            label_visibility="collapsed",
        )
        if multi_arqs:
            adicionados = 0
            for arq in multi_arqs:
                if arq.name not in st.session_state["refs_acumulados"]:
                    st.session_state["refs_acumulados"][arq.name] = arq.getbuffer().tobytes()
                    adicionados += 1
            if adicionados:
                st.success(f"✅ **{adicionados} arquivo(s) novos** adicionados à fila!")

    # ── ABA UM POR VEZ ───────────────────────────────────────────────────────
    with aba_um:
        st.caption("Envie os arquivos um por vez — útil quando os outros métodos não funcionam.")
        novo_arq = st.file_uploader(
            "Selecione um arquivo",
            type=["pdf", "docx", "doc", "txt"],
            key="refs_individual",
            label_visibility="collapsed",
        )
        if novo_arq is not None:
            st.session_state["refs_acumulados"][novo_arq.name] = novo_arq.getbuffer().tobytes()

    # ── FILA ACUMULADA ────────────────────────────────────────────────────────
    total = st.session_state["refs_acumulados"]
    n_total = len(total)
    mb_total = sum(len(b) for b in total.values()) / (1024 * 1024)

    st.markdown("---")
    if n_total > 0:
        st.success(f"✅ **{n_total} arquivo(s)** na fila — {mb_total:.1f} MB total")
        with st.expander(f"Ver lista ({n_total} arquivos)"):
            for nome in list(total.keys()):
                col_nome, col_del = st.columns([5, 1])
                col_nome.markdown(
                    f"📄 {nome} &nbsp;<small>({len(total[nome])//1024} KB)</small>",
                    unsafe_allow_html=True,
                )
                if col_del.button("✕", key=f"del_{nome}", help="Remover"):
                    del st.session_state["refs_acumulados"][nome]
                    st.rerun()
        if st.button("🗑️  Limpar todos os arquivos", use_container_width=True):
            st.session_state["refs_acumulados"] = {}
            st.rerun()
    else:
        st.info(
            "Nenhum arquivo adicionado ainda. "
            "Use uma das abas acima para enviar os arquivos.",
            icon="📂",
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
    refs_acumulados = st.session_state.get("refs_acumulados", {})

    if not docx_file:
        st.error("❌  Selecione o arquivo da dissertação (passo 1).")
        st.stop()
    if not refs_acumulados:
        st.error("❌  Adicione pelo menos um arquivo de referência (passo 2).")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:

        # Salva DOCX
        diss_path = os.path.join(tmpdir, "dissertacao.docx")
        with open(diss_path, "wb") as fh:
            fh.write(docx_file.getbuffer())

        refs_dir = os.path.join(tmpdir, "refs")
        os.makedirs(refs_dir, exist_ok=True)

        # Salva todos os arquivos acumulados
        for nome, conteudo in refs_acumulados.items():
            with open(os.path.join(refs_dir, nome), "wb") as fh:
                fh.write(conteudo)

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
