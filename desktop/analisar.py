#!/usr/bin/env python3
"""
VERIFICADOR DE CITAÇÕES ACADÊMICAS
Versão desktop — roda localmente no seu computador.

Como usar:
  python analisar.py          → abre a interface gráfica
  python analisar.py disser.docx pasta_refs/  → modo terminal

Requisitos mínimos:
  pip install anthropic        (apenas para verificação semântica com IA)
  pip install PyMuPDF          (apenas se tiver referências em PDF)
  pip install python-docx      (fallback para .docx; normalmente não precisa)
"""

import json
import os
import re
import sys
import threading
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

# ═══════════════════════════════════════════════════════════════════════════
# LEITURA DE ARQUIVOS
# ═══════════════════════════════════════════════════════════════════════════

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _xml_para_texto(xml_bytes: bytes) -> str:
    try:
        root = ET.fromstring(xml_bytes)
        linhas = []
        for para in root.iter(f"{{{_NS}}}p"):
            tokens = [t.text for t in para.iter(f"{{{_NS}}}t") if t.text]
            linhas.append("".join(tokens))
        return "\n".join(linhas)
    except Exception:
        return ""


def ler_docx(caminho: Path) -> str:
    partes = []
    try:
        with zipfile.ZipFile(str(caminho)) as z:
            nomes = z.namelist()
            if "word/document.xml" in nomes:
                partes.append(_xml_para_texto(z.read("word/document.xml")))
            if "word/footnotes.xml" in nomes:
                partes.append(_xml_para_texto(z.read("word/footnotes.xml")))
            if "word/endnotes.xml" in nomes:
                partes.append(_xml_para_texto(z.read("word/endnotes.xml")))
    except Exception:
        try:
            import docx as _docx
            doc = _docx.Document(str(caminho))
            partes.append("\n".join(p.text for p in doc.paragraphs))
        except Exception as e:
            return f"[ERRO ao ler .docx: {e}]"
    return re.sub(r"\n{3,}", "\n\n", "\n".join(partes)).strip()


def ler_pdf(caminho: Path) -> str:
    try:
        import fitz
        doc = fitz.open(str(caminho))
        return "\n".join(page.get_text() for page in doc)
    except ImportError:
        return "[PDF: instale PyMuPDF com 'pip install PyMuPDF' para ler PDFs]"
    except Exception as e:
        return f"[ERRO ao ler PDF: {e}]"


def ler_arquivo(caminho: Path) -> str:
    ext = caminho.suffix.lower()
    if ext in (".docx", ".doc"):
        return ler_docx(caminho)
    if ext == ".pdf":
        return ler_pdf(caminho)
    if ext == ".txt":
        try:
            return caminho.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            return f"[ERRO: {e}]"
    return ""


# ═══════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DE CITAÇÕES (padrão ABNT)
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class Citacao:
    texto: str
    autores: list
    ano: str
    pagina: Optional[str]
    contexto: str
    paragrafo_idx: int


@dataclass
class Referencia:
    texto: str
    sobrenome: str
    sobrenomes: list
    ano: str
    titulo: str


_SOBR = r"[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.'){1,}"
_ANO  = r"\d{4}[a-z]?"
_PAG  = r"(?:[,;:\s]+(?:p\.?p?|pág\.?|f\.?)\s*([\d\-–]+))?"
_AUTORES_CAP = (
    r"("
    + _SOBR + r"(?:\s+" + _SOBR + r")?"
    + r"(?:,?\s*et\s+al\.?)?"
    + r"(?:\s*;\s*" + _SOBR + r"(?:\s+" + _SOBR + r")?(?:,?\s*et\s+al\.?)?)*"
    + r")"
)
_RE_PAREN = re.compile(r"\(" + _AUTORES_CAP + r"[,;]\s*(" + _ANO + r")" + _PAG + r"\)")
_RE_NARR  = re.compile(_AUTORES_CAP + r"\s+\((" + _ANO + r")" + _PAG + r"\)")
_RE_INICIO_REFS = re.compile(r"^\s*REFER[EÊ]NCIAS\s*(BIBLIOGR[ÁA]FICAS?)?\s*$", re.IGNORECASE)
_RE_ANO_PLAUSIVEL = re.compile(r"\b((?:19|20)\d{2})\b")

_NAO_SOBRENOME = {
    "PARA", "SEGUNDO", "CONFORME", "SOBRE", "COMO", "POR", "COM",
    "PELO", "PELA", "PELOS", "PELAS", "EM", "DE", "DA", "DO",
    "APUD", "VER", "CF", "VIDE", "ASSIM", "DIANTE",
}


def _autores_paren(raw: str) -> list:
    raw = re.sub(r",?\s*et\s+al\.?", "", raw, flags=re.IGNORECASE)
    result = []
    for p in re.split(r"\s*;\s*", raw):
        tokens = p.strip().split()
        if tokens:
            result.append(tokens[0].upper())
    return result or ["?"]


def _autores_narr(raw: str) -> list:
    raw = re.sub(r",?\s*et\s+al\.?", "", raw, flags=re.IGNORECASE).strip().rstrip(".")
    result = []
    for p in re.split(r"\s*[;,]\s*", raw):
        tokens = [t.strip(".,") for t in p.strip().split()
                  if t.strip(".,").upper() not in _NAO_SOBRENOME]
        if tokens:
            result.append(tokens[-1].upper())
        elif p.strip().split():
            result.append(p.strip().split()[-1].strip(".,").upper())
    return result or ["?"]


def _contexto(paragrafos: list, idx: int, janela: int = 300) -> str:
    antes = " ".join(paragrafos[max(0, idx - 2):idx])
    depois = " ".join(paragrafos[idx + 1:idx + 3])
    return (antes[-janela:] + " " + paragrafos[idx] + " " + depois[:janela]).strip()


def encontrar_citacoes(texto: str):
    paragrafos = texto.split("\n")
    idx_refs = next((i for i, p in enumerate(paragrafos) if _RE_INICIO_REFS.match(p.strip())), -1)
    limite = idx_refs if idx_refs > 0 else len(paragrafos)
    citacoes, vistos = [], set()

    for idx, para in enumerate(paragrafos[:limite]):
        ctx = _contexto(paragrafos[:limite], idx)
        for m in _RE_PAREN.finditer(para):
            aut = _autores_paren(m.group(1) or "")
            ano = m.group(2) or "?"
            pag = m.group(3) if m.lastindex and m.lastindex >= 3 else None
            if (k := (tuple(aut), ano)) not in vistos:
                vistos.add(k)
                citacoes.append(Citacao(m.group(0), aut, ano, pag, ctx, idx))
        for m in _RE_NARR.finditer(para):
            aut = _autores_narr(m.group(1) or "")
            ano = m.group(2) or "?"
            pag = m.group(3) if m.lastindex and m.lastindex >= 3 else None
            if (k := (tuple(aut), ano)) not in vistos:
                vistos.add(k)
                citacoes.append(Citacao(m.group(0), aut, ano, pag, ctx, idx))

    return citacoes, idx_refs


# ═══════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DE REFERÊNCIAS (lista bibliográfica ABNT)
# ═══════════════════════════════════════════════════════════════════════════

def extrair_referencias(texto: str, idx_refs: int) -> list:
    if idx_refs < 0:
        return []
    paragrafos = texto.split("\n")
    referencias, buffer = [], []

    _RE_NOVA = re.compile(
        r"^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+[,;]"
    )
    _RE_AUTORES_REF = re.compile(
        r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\']+),"
    )

    def _sobrenomes_ref(t: str) -> list:
        cab = re.split(r"\.\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][a-záéíóúâêîôûãõàç]", t)[0]
        resultado = [m.group(1).upper() for m in _RE_AUTORES_REF.finditer(cab) if len(m.group(1)) > 1]
        if not resultado:
            m = re.match(r"^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-Za-záéíóúâêîôûãõàç\-\']+)", t)
            if m:
                resultado = [m.group(1).upper()]
        return resultado or ["DESCONHECIDO"]

    def processar(buf: list) -> Optional[Referencia]:
        t = " ".join(buf).strip()
        if not t:
            return None
        sobrenomes = _sobrenomes_ref(t)
        m_ano = _RE_ANO_PLAUSIVEL.search(t)
        ano = m_ano.group(1) if m_ano else "s.d."
        m_tit = re.search(r"\.\s+([A-Z][^.\n]{5,})", t)
        titulo = m_tit.group(1)[:80] if m_tit else t[:80]
        return Referencia(t, sobrenomes[0], sobrenomes, ano, titulo)

    for linha in paragrafos[idx_refs + 1:]:
        linha = linha.strip()
        if not linha:
            if buffer:
                r = processar(buffer)
                if r:
                    referencias.append(r)
                buffer = []
        elif _RE_NOVA.match(linha) and buffer:
            r = processar(buffer)
            if r:
                referencias.append(r)
            buffer = [linha]
        else:
            buffer.append(linha)

    if buffer:
        r = processar(buffer)
        if r:
            referencias.append(r)

    return referencias


# ═══════════════════════════════════════════════════════════════════════════
# CRUZAMENTO: citação ↔ referência
# ═══════════════════════════════════════════════════════════════════════════

def cruzar(citacoes: list, referencias: list) -> dict:
    idx = {}
    for r in referencias:
        for sob in r.sobrenomes:
            chave = f"{sob}_{r.ano}"
            if chave not in idx:
                idx[chave] = r

    chaves_citadas, pareamentos, citadas_sem_ref = set(), [], []

    for c in citacoes:
        encontrou = False
        for sob in c.autores:
            chave = f"{sob}_{c.ano}"
            chaves_citadas.add(chave)
            if chave in idx:
                ref = idx[chave]
                for s in ref.sobrenomes:
                    chaves_citadas.add(f"{s}_{ref.ano}")
                pareamentos.append((c, ref))
                encontrou = True
                break
        if not encontrou:
            citadas_sem_ref.append(c)

    refs_sem_citacao = [r for r in referencias
                        if not any(f"{s}_{r.ano}" in chaves_citadas for s in r.sobrenomes)]

    return {"pareamentos": pareamentos,
            "citadas_sem_ref": citadas_sem_ref,
            "refs_sem_citacao": refs_sem_citacao}


# ═══════════════════════════════════════════════════════════════════════════
# VERIFICAÇÃO SEMÂNTICA COM CLAUDE (opcional)
# ═══════════════════════════════════════════════════════════════════════════

_STOPWORDS = {
    "de", "da", "do", "das", "dos", "e", "a", "o", "as", "os", "um", "uma",
    "que", "em", "para", "com", "por", "se", "na", "no", "nas", "nos",
    "ao", "aos", "pelo", "pela", "pelos", "pelas", "como", "mais", "mas",
    "ou", "sua", "seu", "ser", "foi", "são", "está", "não", "também",
    "quando", "sobre", "entre", "mesmo", "ainda", "pode",
}


def _trecho_relevante(contexto: str, texto_fonte: str, janela: int = 4000) -> str:
    palavras = [p for p in re.findall(r"[a-záéíóúâêîôûãõàç]{4,}", contexto.lower())
                if p not in _STOPWORDS]
    if not palavras:
        return texto_fonte[:janela]
    paras = [p.strip() for p in texto_fonte.split("\n") if len(p.strip()) > 40]
    if not paras:
        return texto_fonte[:janela]
    scores = sorted(range(len(paras)),
                    key=lambda i: -sum(1 for w in palavras if w in paras[i].lower()))
    melhor = scores[0]
    indices = sorted({max(0, melhor - 2), max(0, melhor - 1), melhor,
                      min(len(paras) - 1, melhor + 1), min(len(paras) - 1, melhor + 2)})
    trecho = "\n".join(paras[i] for i in indices)
    return trecho[:janela]


def _buscar_fonte(sobrenome: str, ano: str, titulo: str, material: dict) -> tuple:
    sob_l = sobrenome.lower()
    tit_palavras = [p for p in re.findall(r"[a-záéíóúâêîôûãõàç]{4,}", titulo.lower())
                    if p not in _STOPWORDS][:5]
    candidatos = []
    for nome, texto in material.items():
        score = 0
        tl = texto.lower()
        if sob_l in nome.lower(): score += 10
        if ano in nome:           score += 5
        if sob_l in tl:           score += 3
        if ano in texto:          score += 2
        for p in tit_palavras:
            if p in tl[:2000]:    score += 1
        if score > 0:
            candidatos.append((score, nome, texto))
    if not candidatos:
        return "", ""
    candidatos.sort(key=lambda x: -x[0])
    _, nome_m, texto_m = candidatos[0]
    return nome_m, texto_m


_PROMPT_CLAUDE = """Você é um verificador de integridade acadêmica.

TRECHO DA DISSERTAÇÃO com a citação {citacao}:
\"\"\"{contexto}\"\"\"

REFERÊNCIA BIBLIOGRÁFICA: {referencia}
ARQUIVO FONTE: {arquivo}

TRECHO RELEVANTE DA OBRA CITADA:
\"\"\"{fonte}\"\"\"

Analise:
1. O argumento da dissertação é coerente com o que a obra realmente afirma?
2. Há distorção de sentido, generalização ou uso fora de contexto?
3. A ideia atribuída ao autor consta na obra?

Responda SOMENTE com JSON (sem markdown):
{{"veredicto": "CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE", "justificativa": "frase curta", "trecho_fonte": "trecho real da obra até 150 chars"}}
"""


def verificar_claude(cit: Citacao, ref: Referencia, arquivo: str,
                     texto_fonte: str, api_key: str) -> dict:
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        trecho = _trecho_relevante(cit.contexto, texto_fonte)
        prompt = _PROMPT_CLAUDE.format(
            citacao=cit.texto,
            contexto=cit.contexto[:600],
            referencia=ref.texto[:150],
            arquivo=arquivo,
            fonte=trecho,
        )
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = re.sub(r"^```(?:json)?", "", resp.content[0].text.strip()).strip("`").strip()
        try:
            return json.loads(raw)
        except Exception:
            m = re.search(r"\{[^}]+\}", raw, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group())
                except Exception:
                    pass
        return {"veredicto": "ERRO", "justificativa": raw[:200], "trecho_fonte": ""}
    except Exception as e:
        return {"veredicto": "ERRO", "justificativa": str(e)[:150], "trecho_fonte": ""}


# ═══════════════════════════════════════════════════════════════════════════
# RELATÓRIO HTML
# ═══════════════════════════════════════════════════════════════════════════

def gerar_html(resultado: dict) -> str:
    ts          = datetime.now().strftime("%Y-%m-%d %H:%M")
    diss_path   = resultado["diss_path"]
    citadas_sr  = resultado["citadas_sem_ref"]
    refs_sc     = resultado["refs_sem_citacao"]
    pareamentos = resultado["pareamentos"]
    verifics    = resultado["verificacoes"]

    corretas   = [v for v in verifics if v.get("veredicto") == "CORRETO"]
    incorretas = [v for v in verifics if v.get("veredicto") == "INCORRETO"]
    parciais   = [v for v in verifics if v.get("veredicto") == "PARCIAL"]
    sem_fonte  = [v for v in verifics if v.get("veredicto") == "SEM_FONTE"]

    def e(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def secao(titulo, items, cor, fn):
        if not items:
            return ""
        rows = "".join(f"<li>{fn(i)}</li>" for i in items)
        return f'<h2 style="color:{cor}">{e(titulo)} ({len(items)})</h2><ul>{rows}</ul>'

    def card(v):
        t = v.get("trecho_fonte", "")
        bq = f'<blockquote>"{e(t)}"</blockquote>' if t else ""
        return (f'<b>{e(v["citacao"])}</b>'
                f'<br><small>Referência: {e(v.get("referencia","")[:150])}</small>'
                f'<br><small>Arquivo fonte: <em>{e(v.get("arquivo_fonte","—"))}</em></small>'
                f'<br>{e(v.get("justificativa",""))}{bq}'
                f'<details><summary><small>ver contexto</small></summary>'
                f'<small>{e(v.get("contexto","")[:300])}</small></details>')

    n_cit = len(resultado["citacoes"])
    n_ref = len(resultado["referencias"])
    n_par = len(pareamentos)
    n_csr = len(citadas_sr)
    n_rsc = len(refs_sc)
    n_ok  = len(corretas)
    n_par2 = len(parciais)
    n_err = len(incorretas)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>Verificador de Citações</title>
<style>
* {{box-sizing:border-box}}
body {{font-family:Segoe UI,Arial,sans-serif;max-width:980px;margin:2em auto;padding:0 1.5em;line-height:1.6;color:#222}}
h1 {{background:#1a2744;color:#fff;padding:.9em 1.2em;border-radius:8px;margin-bottom:.5em;font-size:1.4em}}
h2 {{margin-top:2em;border-bottom:3px solid currentColor;padding-bottom:.3em;font-size:1.15em}}
ul {{list-style:none;padding:0}}
li {{border-left:5px solid #ccc;padding:.7em 1em;margin:.5em 0;background:#fafafa;border-radius:0 6px 6px 0}}
.ok {{border-color:#27ae60;background:#f0fff4}}
.err {{border-color:#e74c3c;background:#fff5f5}}
.warn {{border-color:#f39c12;background:#fffbf0}}
.info {{border-color:#3498db;background:#f0f8ff}}
.resumo {{display:flex;gap:1em;flex-wrap:wrap;margin:1.2em 0}}
.card {{border-radius:10px;padding:1em 1.5em;min-width:120px;text-align:center;border-top:5px solid #ccc}}
.card b {{display:block;font-size:2.5em;line-height:1}}
blockquote {{margin:.5em 0;padding:.5em 1em;background:#eef3ff;border-left:4px solid #6c8ebf;font-style:italic;color:#333;border-radius:0 4px 4px 0}}
details summary {{cursor:pointer;color:#666;font-size:.85em;margin-top:.3em}}
small {{color:#555;font-size:.85em}}
</style></head><body>
<h1>Verificador de Citações Acadêmicas</h1>
<p>Gerado em <b>{e(ts)}</b> | Arquivo: <em>{e(diss_path)}</em></p>
<div class="resumo">
  <div class="card" style="border-color:#3498db"><b>{n_cit}</b>Citações<br>encontradas</div>
  <div class="card" style="border-color:#8e44ad"><b>{n_ref}</b>Referências<br>na lista</div>
  <div class="card" style="border-color:#27ae60"><b>{n_par}</b>Pares<br>citação-ref</div>
  <div class="card" style="border-color:#e74c3c"><b>{n_csr}</b>Cit. sem<br>referência</div>
  <div class="card" style="border-color:#e67e22"><b>{n_rsc}</b>Ref. sem<br>citação</div>
  <div class="card" style="border-color:#27ae60"><b>{n_ok}</b>Verificadas<br>OK</div>
  <div class="card" style="border-color:#f39c12"><b>{n_par2}</b>Parciais</div>
  <div class="card" style="border-color:#e74c3c"><b>{n_err}</b>Incorretas</div>
</div>
"""

    html += secao("Citacoes SEM referencia na lista bibliografica", citadas_sr, "#e74c3c",
        lambda c: (f'<b class="err">{e(c.texto)}</b>'
                   f'<details><summary><small>ver contexto</small></summary>'
                   f'<small>{e(c.contexto[:350])}</small></details>'))

    html += secao("Referencias listadas SEM citacao no texto", refs_sc, "#e67e22",
        lambda r: f'<span class="warn">{e(r.texto[:250])}</span>')

    if incorretas:
        html += secao("INCORRETAS - argumento diverge da obra", incorretas, "#c0392b", card)
    if parciais:
        html += secao("PARCIAIS - citacao com ressalvas", parciais, "#d35400", card)
    if corretas:
        html += secao("CORRETAS - confirmadas pela obra", corretas, "#27ae60", card)
    if sem_fonte:
        html += secao("SEM FONTE - arquivo nao encontrado na pasta", sem_fonte, "#7f8c8d", card)

    if not verifics and pareamentos:
        html += f'<h2 style="color:#3498db">Citacoes pareadas com referencias ({n_par})</h2>'
        html += '<p><em>Sem verificacao semantica - informe a chave de API Claude para verificar a coerencia dos argumentos.</em></p><ul>'
        for c, r in pareamentos:
            html += (f'<li class="info"><b>{e(c.texto)}</b>'
                     f'<br><small>Ref.: {e(r.texto[:180])}</small>'
                     f'<details><summary><small>ver contexto</small></summary>'
                     f'<small>{e(c.contexto[:300])}</small></details></li>')
        html += "</ul>"

    html += "</body></html>"
    return html


# ═══════════════════════════════════════════════════════════════════════════
# NÚCLEO DA ANÁLISE
# ═══════════════════════════════════════════════════════════════════════════

def analisar(diss_path: str, refs_dir: str, api_key: str, log_fn, stop_fn=None) -> Optional[dict]:
    def L(msg):
        log_fn(msg)

    L("=" * 65)
    L("ETAPA 1 - Lendo a dissertacao")
    L("=" * 65)
    texto = ler_docx(Path(diss_path))
    L(f"  {len(texto):,} caracteres extraidos")
    if len(texto) < 300:
        L("ERRO: texto muito curto! O arquivo pode ser uma imagem escaneada.")
        L(f"  Conteudo encontrado: {repr(texto[:200])}")
        return None

    L("  Primeiras linhas:")
    for linha in texto.split("\n")[:5]:
        if linha.strip():
            L(f"  | {linha[:100]}")

    L("")
    L("  Identificando citacoes e referencias na dissertacao...")
    citacoes, idx_refs = encontrar_citacoes(texto)
    L(f"  {len(citacoes)} citacao(oes) encontrada(s)")
    if idx_refs >= 0:
        L(f"  Secao REFERENCIAS encontrada (linha {idx_refs})")
    else:
        L("  ATENCAO: Secao REFERENCIAS nao encontrada no arquivo!")

    referencias = extrair_referencias(texto, idx_refs)
    L(f"  {len(referencias)} entrada(s) na lista de referencias")

    L("")
    L("=" * 65)
    L("ETAPA 2 - Lendo arquivos de apoio (um a um)")
    L("=" * 65)
    extensoes = {".pdf", ".docx", ".doc", ".txt"}
    arquivos = sorted(
        [f for f in Path(refs_dir).rglob("*") if f.is_file() and f.suffix.lower() in extensoes],
        key=lambda f: f.name.lower()
    )
    L(f"  {len(arquivos)} arquivo(s) encontrado(s) na pasta")
    L("")

    material = {}
    for i, arq in enumerate(arquivos, 1):
        if stop_fn and stop_fn():
            L("\n[INTERROMPIDO]")
            return None
        t = ler_arquivo(arq)
        ok = t and len(t) > 50 and not t.startswith("[")
        if ok:
            material[arq.name] = t
            L(f"  [{i:>3}/{len(arquivos)}]  {arq.name}  ({len(t):,} chars)")
        else:
            L(f"  [{i:>3}/{len(arquivos)}]  {arq.name}  -- falha na leitura")

    L(f"\n  {len(material)} arquivo(s) lido(s) com sucesso")

    L("")
    L("=" * 65)
    L("ETAPA 3 - Cruzando citacoes com referencias")
    L("=" * 65)
    cruzamento = cruzar(citacoes, referencias)
    pareamentos      = cruzamento["pareamentos"]
    citadas_sem_ref  = cruzamento["citadas_sem_ref"]
    refs_sem_citacao = cruzamento["refs_sem_citacao"]
    L(f"  Pares encontrados:          {len(pareamentos)}")
    L(f"  Citadas SEM referencia:     {len(citadas_sem_ref)}")
    L(f"  Referencias SEM citacao:    {len(refs_sem_citacao)}")

    if citadas_sem_ref:
        L("\n  Citacoes sem referencia:")
        for c in citadas_sem_ref[:10]:
            L(f"    -> {c.texto}")
        if len(citadas_sem_ref) > 10:
            L(f"    ... e mais {len(citadas_sem_ref) - 10}")

    verificacoes = []
    if api_key.strip() and material and pareamentos:
        L("")
        L("=" * 65)
        L(f"ETAPA 4 - Verificando coerencia argumentativa ({len(pareamentos)} citacoes)")
        L("=" * 65)
        for i, (cit, ref) in enumerate(pareamentos, 1):
            if stop_fn and stop_fn():
                L("\n[INTERROMPIDO]")
                break
            arq, txt = _buscar_fonte(cit.autores[0], cit.ano, ref.titulo, material)
            L(f"  [{i:>3}/{len(pareamentos)}] {cit.texto[:70]}")
            if not txt:
                L(f"         -> arquivo nao encontrado para {ref.sobrenome} ({ref.ano})")
                verificacoes.append({
                    "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                    "contexto": cit.contexto, "referencia": ref.texto,
                    "arquivo_fonte": "", "veredicto": "SEM_FONTE",
                    "justificativa": f"Arquivo de {ref.sobrenome} ({ref.ano}) nao encontrado na pasta",
                    "trecho_fonte": "",
                })
                continue
            v = verificar_claude(cit, ref, arq, txt, api_key)
            emoji = {"CORRETO": "OK", "INCORRETO": "ERRO", "PARCIAL": "PARCIAL",
                     "SEM_FONTE": "?", "ERRO": "!"}.get(v.get("veredicto", ""), "?")
            L(f"         {emoji} {v.get('veredicto')} -- {v.get('justificativa','')[:80]}")
            verificacoes.append({
                "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                "contexto": cit.contexto, "referencia": ref.texto,
                "arquivo_fonte": arq, **v,
            })
    elif not api_key.strip():
        L("")
        L("  [Sem chave API -- verificacao semantica nao realizada]")
        L("  [Para verificar a coerencia dos argumentos, informe a chave Anthropic]")

    return {
        "citacoes":        citacoes,
        "referencias":     referencias,
        "pareamentos":     pareamentos,
        "citadas_sem_ref": citadas_sem_ref,
        "refs_sem_citacao":refs_sem_citacao,
        "verificacoes":    verificacoes,
        "diss_path":       diss_path,
    }


# ═══════════════════════════════════════════════════════════════════════════
# INTERFACE GRÁFICA (tkinter)
# ═══════════════════════════════════════════════════════════════════════════

def _iniciar_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    _parar = threading.Event()
    _resultado = {}

    janela = tk.Tk()
    janela.title("Verificador de Citacoes Academicas")
    janela.geometry("820x680")
    janela.resizable(True, True)

    topo = tk.Frame(janela, bg="#1a2744", padx=12, pady=10)
    topo.pack(fill="x")
    tk.Label(topo, text="Verificador de Citacoes Academicas",
             bg="#1a2744", fg="white", font=("Arial", 15, "bold")).pack(anchor="w")
    tk.Label(topo, text="Le a dissertacao · Le os arquivos de apoio · Cruza e verifica",
             bg="#1a2744", fg="#aac", font=("Arial", 10)).pack(anchor="w")

    form = tk.Frame(janela, padx=16, pady=10)
    form.pack(fill="x")

    var_diss = tk.StringVar()
    var_refs = tk.StringVar()
    var_api  = tk.StringVar()
    var_sem  = tk.BooleanVar(value=False)

    def _campo(parent, texto, var, comando, row):
        tk.Label(parent, text=texto, anchor="w", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=4)
        tk.Entry(parent, textvariable=var, width=58, font=("Arial", 10)).grid(
            row=row, column=1, padx=(6, 4), pady=4, sticky="ew")
        tk.Button(parent, text="Procurar...", command=comando,
                  font=("Arial", 9)).grid(row=row, column=2, pady=4)

    def escolher_diss():
        p = filedialog.askopenfilename(
            title="Selecione a dissertacao",
            filetypes=[("Word", "*.docx *.doc"), ("Todos", "*.*")])
        if p:
            var_diss.set(p)

    def escolher_refs():
        p = filedialog.askdirectory(title="Selecione a pasta com os arquivos de referencia")
        if p:
            var_refs.set(p)

    form.columnconfigure(1, weight=1)
    _campo(form, "Dissertacao (.docx):", var_diss, escolher_diss, 0)
    _campo(form, "Pasta de referencias:", var_refs, escolher_refs, 1)

    tk.Label(form, text="Chave API Claude\n(opcional):", anchor="w",
             font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=4)
    e_api = tk.Entry(form, textvariable=var_api, width=58, font=("Arial", 10), show="*")
    e_api.grid(row=2, column=1, padx=(6, 4), pady=4, sticky="ew")
    tk.Button(form, text="Ver/Ocultar",
              command=lambda: e_api.config(show="" if e_api.cget("show") == "*" else "*"),
              font=("Arial", 9)).grid(row=2, column=2, pady=4)

    tk.Checkbutton(form, text="Apenas cruzamento (sem verificacao IA, mais rapido)",
                   variable=var_sem, font=("Arial", 10)).grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

    btn_frame = tk.Frame(janela, padx=16)
    btn_frame.pack(fill="x")

    btn_iniciar = tk.Button(btn_frame, text="Iniciar analise",
                            font=("Arial", 11, "bold"), bg="#27ae60", fg="white",
                            padx=14, pady=6)
    btn_parar   = tk.Button(btn_frame, text="Parar",
                            font=("Arial", 11), bg="#e74c3c", fg="white",
                            padx=14, pady=6, state="disabled")
    btn_relat   = tk.Button(btn_frame, text="Abrir relatorio HTML",
                            font=("Arial", 11), bg="#2980b9", fg="white",
                            padx=14, pady=6, state="disabled")

    btn_iniciar.pack(side="left", padx=(0, 8), pady=8)
    btn_parar.pack(side="left", padx=(0, 8), pady=8)
    btn_relat.pack(side="left", pady=8)

    prog_var = tk.DoubleVar()
    prog_bar = ttk.Progressbar(janela, variable=prog_var, mode="indeterminate")
    prog_bar.pack(fill="x", padx=16, pady=(0, 4))

    log_area = scrolledtext.ScrolledText(janela, font=("Consolas", 9),
                                         bg="#1e1e1e", fg="#d4d4d4",
                                         insertbackground="white", wrap="word")
    log_area.pack(fill="both", expand=True, padx=16, pady=(0, 12))

    def log(msg):
        def _():
            log_area.configure(state="normal")
            log_area.insert("end", msg + "\n")
            log_area.see("end")
            log_area.configure(state="disabled")
        janela.after(0, _)

    def abrir_relatorio():
        if _resultado.get("html_path"):
            webbrowser.open(f"file:///{_resultado['html_path']}")

    btn_relat.configure(command=abrir_relatorio)

    def iniciar():
        diss = var_diss.get().strip()
        refs = var_refs.get().strip()
        if not diss or not Path(diss).exists():
            messagebox.showerror("Erro", "Selecione um arquivo de dissertacao valido.")
            return
        if not refs or not Path(refs).is_dir():
            messagebox.showerror("Erro", "Selecione uma pasta de referencias valida.")
            return

        _parar.clear()
        btn_iniciar.configure(state="disabled")
        btn_parar.configure(state="normal")
        btn_relat.configure(state="disabled")
        log_area.configure(state="normal")
        log_area.delete("1.0", "end")
        log_area.configure(state="disabled")
        prog_bar.start(12)

        def rodar():
            try:
                resultado = analisar(
                    diss_path=diss,
                    refs_dir=refs,
                    api_key=var_api.get().strip(),
                    log_fn=log,
                    stop_fn=_parar.is_set,
                )
                if resultado:
                    html_str = gerar_html(resultado)
                    saida = Path(diss).parent / "relatorio_citacoes.html"
                    saida.write_text(html_str, encoding="utf-8")
                    _resultado["html_path"] = str(saida).replace("\\", "/")
                    log("")
                    log("=" * 65)
                    log("ANALISE CONCLUIDA!")
                    log(f"  Relatorio salvo em: {saida}")
                    log("=" * 65)
                    def _ativar():
                        btn_relat.configure(state="normal")
                        prog_bar.stop()
                        prog_var.set(100)
                        messagebox.showinfo("Concluido",
                            f"Analise finalizada!\n\n"
                            f"- {len(resultado['citacoes'])} citacoes encontradas\n"
                            f"- {len(resultado['referencias'])} referencias na lista\n"
                            f"- {len(resultado['pareamentos'])} pares citacao/referencia\n"
                            f"- {len(resultado['citadas_sem_ref'])} citadas sem referencia\n"
                            f"- {len(resultado['refs_sem_citacao'])} referencias sem citacao\n\n"
                            f"Relatorio salvo em:\n{saida}")
                    janela.after(0, _ativar)
                else:
                    log("\n[Analise nao concluida -- verifique os erros acima]")
            except Exception as exc:
                import traceback
                log(f"\n[ERRO INESPERADO]\n{traceback.format_exc()}")
            finally:
                def _reset():
                    btn_iniciar.configure(state="normal")
                    btn_parar.configure(state="disabled")
                    prog_bar.stop()
                janela.after(0, _reset)

        threading.Thread(target=rodar, daemon=True).start()

    def parar():
        _parar.set()
        btn_parar.configure(state="disabled")
        log("\n[Solicitando parada...]")

    btn_iniciar.configure(command=iniciar)
    btn_parar.configure(command=parar)

    log("Bem-vindo ao Verificador de Citacoes Academicas.")
    log("1. Selecione o arquivo da dissertacao (.docx)")
    log("2. Selecione a pasta com os arquivos de referencia (PDF, DOCX, TXT)")
    log("3. Informe a chave API Claude (opcional) para verificacao de coerencia")
    log("4. Clique em Iniciar analise")
    log("")

    janela.mainloop()


# ═══════════════════════════════════════════════════════════════════════════
# MODO TERMINAL (sem GUI)
# ═══════════════════════════════════════════════════════════════════════════

def _modo_terminal(diss_path: str, refs_dir: str, api_key: str = ""):
    resultado = analisar(diss_path, refs_dir, api_key, print)
    if not resultado:
        sys.exit(1)
    html_str = gerar_html(resultado)
    saida = Path(diss_path).parent / "relatorio_citacoes.html"
    saida.write_text(html_str, encoding="utf-8")
    print(f"\nRelatorio salvo: {saida}")
    webbrowser.open(f"file:///{saida}")


# ═══════════════════════════════════════════════════════════════════════════
# ENTRYPOINT
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        _modo_terminal(
            sys.argv[1],
            sys.argv[2],
            sys.argv[3] if len(sys.argv) > 3 else "",
        )
    else:
        try:
            _iniciar_gui()
        except ImportError:
            print("tkinter nao disponivel. Use: python analisar.py dissertacao.docx pasta_refs/")
            sys.exit(1)
