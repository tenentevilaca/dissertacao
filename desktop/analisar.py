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

import base64
import json
import os
import re
import sys
import threading
import time
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


_SOBR = r"[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,}"
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


def _split_citacoes_combinadas(para: str) -> str:
    """
    Desdobra citações combinadas num mesmo parêntese em citações separadas.
    Ex: (BRASIL, 2018; FRAZÃO, 2018) → (BRASIL, 2018) (FRAZÃO, 2018)
    Não altera (CEPIK; BORBA, 2011) pois não há ano antes do ponto-e-vírgula.
    """
    def split_paren(m):
        interior = m.group(1)
        if not re.search(r'\d{4}[a-z]?\s*;\s*[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]', interior):
            return m.group(0)
        partes = re.split(r'\s*;\s*(?=[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ])', interior)
        return " ".join(f"({p.strip()})" for p in partes if p.strip())
    return re.sub(r'\(([^)]+)\)', split_paren, para)


def encontrar_citacoes(texto: str, log_fn=None):
    def _log(msg):
        if log_fn:
            log_fn(msg)

    paragrafos = texto.split("\n")
    total_paras = len(paragrafos)

    # Use the LAST occurrence of the REFERÊNCIAS heading — TOC entries appear
    # at the top and would give a near-zero scanning window if taken first.
    ocorrencias = [i for i, p in enumerate(paragrafos) if _RE_INICIO_REFS.match(p.strip())]
    if ocorrencias:
        # If multiple matches, prefer one in the latter half of the document
        metade = total_paras // 2
        tardios = [i for i in ocorrencias if i >= metade]
        idx_refs = tardios[-1] if tardios else ocorrencias[-1]
    else:
        idx_refs = -1

    limite = idx_refs if idx_refs > 0 else total_paras
    _log(f"  [DIAG] Total parágrafos: {total_paras} | "
         f"idx_refs={idx_refs} | varredura: parágrafos 0–{limite-1}")
    if ocorrencias and len(ocorrencias) > 1:
        _log(f"  [DIAG] 'REFERÊNCIAS' encontrado {len(ocorrencias)}× nas linhas: "
             f"{ocorrencias} — usando última ocorrência após a metade ({idx_refs})")

    citacoes, vistos = [], set()

    for idx, para in enumerate(paragrafos[:limite]):
        ctx = _contexto(paragrafos[:limite], idx)
        para_exp = _split_citacoes_combinadas(para)
        for m in _RE_PAREN.finditer(para_exp):
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

    if not citacoes and limite > 50:
        com_parens = sum(1 for p in paragrafos[:limite] if '(' in p and ')' in p)
        _log(f"  [DIAG] Parágrafos com parênteses na área varrida: {com_parens} de {limite}")
        _log("  [DIAG] Amostra dos primeiros 20 parágrafos não-vazios:")
        conta = 0
        for p in paragrafos[:limite]:
            if p.strip() and conta < 20:
                _log(f"    | {p[:120]}")
                conta += 1

    return citacoes, idx_refs


# ═══════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DE REFERÊNCIAS (lista bibliográfica ABNT)
# ═══════════════════════════════════════════════════════════════════════════

def extrair_referencias(texto: str, idx_refs: int, log_fn=None) -> list:
    def _log(msg):
        if log_fn:
            log_fn(msg)

    if idx_refs < 0:
        return []
    paragrafos = texto.split("\n")
    referencias, buffer = [], []

    _RE_FIM_REFS = re.compile(
        r"^\s*(AP[EÊ]NDICE[S]?|ANEXO[S]?|GLOSS[ÁA]RIO[S]?|QUESTION[ÁA]RIO[S]?)\b",
        re.IGNORECASE
    )

    # 1ª alt: SOBRENOME, (nome pessoal)
    # 2ª alt: BRASIL. / OSCE. (institucional de 1 palavra seguida de ponto)
    # 3ª alt: ORGANIZAÇÃO PARA A... (SIGLA). ou FÓRUM BRASILEIRO... (institucional multi-palavra)
    _RE_NOVA = re.compile(
        r"^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+[,;]"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,}[a-záéíóúâêîôûãõàç\'\-]*\.\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,}(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç]*)+\s*(?:\([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,12}\)\s*)?[,.]"
    )
    _RE_AUTORES_REF = re.compile(
        r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\']+),"
    )

    def _sobrenomes_ref(t: str) -> list:
        cab = re.split(r"\.\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][a-záéíóúâêîôûãõàç]", t)[0]
        resultado = [m.group(1).upper() for m in _RE_AUTORES_REF.finditer(cab)
                     if len(m.group(1)) > 1]
        if not resultado:
            # Usa classe completa de acentuados para não truncar "ORGANIZAÇÃO" → "ORGANIZA"
            m = re.match(
                r"^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\']+)", t
            )
            if m:
                resultado = [m.group(1).upper()]
        # Extrai siglas entre parênteses: "FÓRUM BRASILEIRO (FBSP)." → indexa FBSP
        siglas = re.findall(r'\(([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,12})\)', t[:300])
        _excluir = {'ORG', 'ED', 'COORD', 'TRAD', 'COMP', 'REV'}
        for s in siglas:
            if s not in resultado and s not in _excluir:
                resultado.append(s)
        return resultado or ["DESCONHECIDO"]

    def processar(buf: list) -> Optional[Referencia]:
        t = " ".join(buf).strip()
        if not t:
            return None
        sobrenomes = _sobrenomes_ref(t)
        m_ano = _RE_ANO_PLAUSIVEL.search(t)
        ano = m_ano.group(1) if m_ano else "s.d."
        m_tit = re.search(r"\.\s+([A-Z][^\.\n]{5,})", t)
        titulo = m_tit.group(1)[:80] if m_tit else t[:80]
        r = Referencia(t, sobrenomes[0], sobrenomes, ano, titulo)
        _log(f"    [REF] {sobrenomes} | {ano} | {t[:80]}")
        return r

    for linha in paragrafos[idx_refs + 1:]:
        linha = linha.strip()
        # Para ao encontrar início de apêndice/anexo/glossário
        if _RE_FIM_REFS.match(linha):
            if buffer:
                r = processar(buffer)
                if r:
                    referencias.append(r)
            break
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

import unicodedata as _ucd

def _normalizar(s: str) -> str:
    return _ucd.normalize("NFD", s.lower()).encode("ascii", "ignore").decode()

def _editar_dist(a: str, b: str) -> int:
    """Distância de edição (Levenshtein) entre duas strings."""
    if abs(len(a) - len(b)) > 2:
        return 99
    dp = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        prev, dp[0] = dp[0], i
        for j, cb in enumerate(b, 1):
            prev, dp[j] = dp[j], prev if ca == cb else 1 + min(prev, dp[j], dp[j-1])
    return dp[len(b)]


def _fuzzy_sobrenome(a: str, b: str) -> bool:
    """True se os dois sobrenomes são suficientemente similares (tolerância a acentos e grafias)."""
    na, nb = _normalizar(a), _normalizar(b)
    if na == nb:
        return True
    # Um é prefixo significativo do outro (FERREIRA ↔ FERREIR)
    curto, longo = (na, nb) if len(na) <= len(nb) else (nb, na)
    if len(curto) >= 5 and longo.startswith(curto):
        return True
    # Distância de edição ≤ 1 para nomes ≥ 5 chars (SOUSA ↔ SOUZA, NÓBREGA ↔ NOBREGA)
    if len(na) >= 5 and len(nb) >= 5 and _editar_dist(na, nb) <= 1:
        return True
    # Jaccard de bigramas ≥ 0.72 para nomes mais longos (≥ 7 chars)
    if len(na) >= 7 and len(nb) >= 7:
        def bg(s): return {s[i:i+2] for i in range(len(s)-1)}
        bga, bgb = bg(na), bg(nb)
        jaccard = len(bga & bgb) / len(bga | bgb)
        if jaccard >= 0.72:
            return True
    return False


def cruzar(citacoes: list, referencias: list, log_fn=None) -> dict:
    def _log(msg):
        if log_fn:
            log_fn(msg)

    # Índice exato: "SOBRENOME_ANO" → referência
    idx_exato: dict = {}
    for r in referencias:
        for sob in r.sobrenomes:
            chave = f"{sob.upper()}_{r.ano}"
            if chave not in idx_exato:
                idx_exato[chave] = r

    chaves_citadas: set = set()
    pareamentos, citadas_sem_ref = [], []
    ja_pareados: set = set()

    for c in citacoes:
        encontrou = False
        for sob in c.autores:
            chave = f"{sob.upper()}_{c.ano}"
            chaves_citadas.add(chave)

            # 1) Busca exata (sobrenome + ano idênticos)
            if chave in idx_exato:
                ref = idx_exato[chave]
                for s in ref.sobrenomes:
                    chaves_citadas.add(f"{s.upper()}_{ref.ano}")
                par_id = id(ref)
                if par_id not in ja_pareados:
                    ja_pareados.add(par_id)
                    pareamentos.append((c, ref))
                encontrou = True
                break

            # 2) Fuzzy: sobrenome similar + mesmo ano
            if not encontrou:
                for r in referencias:
                    if r.ano == c.ano and any(_fuzzy_sobrenome(sob, rs) for rs in r.sobrenomes):
                        for s in r.sobrenomes:
                            chaves_citadas.add(f"{s.upper()}_{r.ano}")
                        par_id = id(r)
                        if par_id not in ja_pareados:
                            ja_pareados.add(par_id)
                            pareamentos.append((c, r))
                        encontrou = True
                        break

            # 3) Fuzzy tolerante: sobrenome similar + ano ±1
            if not encontrou:
                try:
                    ano_cit = int(c.ano[:4])
                    for r in referencias:
                        try:
                            ano_ref = int(r.ano[:4])
                        except ValueError:
                            continue
                        if (abs(ano_cit - ano_ref) <= 1
                                and any(_fuzzy_sobrenome(sob, rs) for rs in r.sobrenomes)):
                            for s in r.sobrenomes:
                                chaves_citadas.add(f"{s.upper()}_{r.ano}")
                            par_id = id(r)
                            if par_id not in ja_pareados:
                                ja_pareados.add(par_id)
                                pareamentos.append((c, r))
                            encontrou = True
                            break
                except (ValueError, TypeError):
                    pass

        if not encontrou:
            citadas_sem_ref.append(c)
            # Diagnóstico: mostra o que foi buscado e os candidatos mais próximos
            buscados = [f"{s.upper()}_{c.ano}" for s in c.autores]
            _log(f"    [SEM-PAR] {c.texto} | buscou: {buscados}")
            candidatos = []
            for r in referencias:
                for rs in r.sobrenomes:
                    for sob in c.autores:
                        if _fuzzy_sobrenome(sob, rs):
                            candidatos.append(f"{rs}_{r.ano}")
            if candidatos:
                _log(f"             candidatos fuzzy: {candidatos[:5]}")
            else:
                _log(f"             nenhum candidato fuzzy encontrado nas {len(referencias)} referências")

    refs_sem_citacao = [r for r in referencias
                        if not any(f"{s.upper()}_{r.ano}" in chaves_citadas
                                   for s in r.sobrenomes)]

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


def _trecho_relevante(contexto: str, texto_fonte: str, janela: int = 20000) -> str:
    """Seleciona os trechos mais relevantes do texto fonte para enviar à IA.
    Envia até `janela` caracteres (~10 páginas) distribuídos pelos parágrafos
    com maior sobreposição de palavras-chave com o contexto da citação.
    """
    palavras = [p for p in re.findall(r"[a-záéíóúâêîôûãõàç]{4,}", contexto.lower())
                if p not in _STOPWORDS]
    if not palavras:
        return texto_fonte[:janela]
    paras = [p.strip() for p in texto_fonte.split("\n") if len(p.strip()) > 40]
    if not paras:
        return texto_fonte[:janela]

    # Pontua cada parágrafo pela quantidade de palavras-chave presentes
    scored = sorted(
        range(len(paras)),
        key=lambda i: -sum(1 for w in palavras if w in paras[i].lower())
    )

    # Inclui os 3 melhores parágrafo com ±3 vizinhos cada para dar contexto
    indices = set()
    for rank in scored[:3]:
        for delta in range(-3, 4):
            idx = rank + delta
            if 0 <= idx < len(paras):
                indices.add(idx)

    trecho = "\n".join(paras[i] for i in sorted(indices))

    # Se ainda tiver espaço na janela, acrescenta o início da obra (capa/intro)
    if len(trecho) < janela // 2:
        inicio = "\n".join(paras[:10])
        trecho = inicio + "\n\n[...]\n\n" + trecho

    return trecho[:janela]


def _extrair_palavras_titulo(ref_texto: str) -> tuple:
    """
    Extrai palavras-chave do título a partir do texto completo da referência.
    Retorna (palavras_do_artigo/capitulo, palavras_do_livro).
    Detecta compêndios pelo padrão ABNT: "... In: EDITOR (org.)..."
    """
    palavras_artigo = []
    palavras_livro  = []

    # Detecta compêndio (capítulo de livro): "Título do capítulo. In: EDITOR..."
    m_in = re.search(r"\bIn:\s*", ref_texto, re.IGNORECASE)
    if m_in:
        texto_cap  = ref_texto[:m_in.start()]
        texto_livro = ref_texto[m_in.end():]
        palavras_artigo = [p for p in re.findall(r"[a-záéíóúâêîôûãõàçA-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{5,}",
                                                  texto_cap) if p.lower() not in _STOPWORDS][:12]
        palavras_livro  = [p for p in re.findall(r"[a-záéíóúâêîôûãõàçA-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{5,}",
                                                  texto_livro) if p.lower() not in _STOPWORDS][:12]
    else:
        # Referência simples: pega tudo após o primeiro "."
        m_titulo = re.search(r"\.\s+(.+)", ref_texto)
        src = m_titulo.group(1) if m_titulo else ref_texto
        palavras_artigo = [p for p in re.findall(r"[a-záéíóúâêîôûãõàçA-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{5,}",
                                                   src) if p.lower() not in _STOPWORDS][:15]

    return (
        [p.lower() for p in palavras_artigo],
        [p.lower() for p in palavras_livro],
    )


def _buscar_fonte(sobrenome: str, ano: str, ref_texto: str,
                  material: dict, log_fn=None) -> tuple:
    """
    Encontra o arquivo de apoio para uma referência bibliográfica.

    Lógica de scoring (maior = melhor):
    ─ Sobrenome no nome do arquivo      → +20
    ─ Ano no nome do arquivo            → +10
    ─ Palavras do título no conteúdo    → +2 por palavra (cap) / +1 por palavra (livro)
    ─ Sobrenome no conteúdo             → +5 (se já no filename) / +3 (só conteúdo)
    ─ Ano no conteúdo                   → +3 (se já no filename) / +1 (só conteúdo)

    Para compêndios (detectados por "In:"), busca separadamente pelo
    título do capítulo E pelo título do livro, e aceita o arquivo que
    contenha o autor do capítulo OU o título do livro.
    """
    sob_l = sobrenome.lower()
    palavras_cap, palavras_livro = _extrair_palavras_titulo(ref_texto)
    eh_compendio = bool(palavras_livro)

    candidatos = []
    for nome, texto in material.items():
        score = 0
        nome_l = nome.lower()
        tl = texto.lower() if texto else ""

        # ── Matching por nome de arquivo ──────────────────────────────
        nome_no_arquivo = sob_l in nome_l
        ano_no_arquivo  = ano in nome
        if nome_no_arquivo: score += 20
        if ano_no_arquivo:  score += 10

        # ── Matching por conteúdo ─────────────────────────────────────
        if tl:
            sob_no_texto = sob_l in tl
            ano_no_texto = ano in tl

            # Palavras do título do artigo/capítulo
            hits_cap = sum(1 for p in palavras_cap if p in tl)
            # Palavras do título do livro (compêndio)
            hits_livro = sum(1 for p in palavras_livro if p in tl)

            if nome_no_arquivo:
                # Filename já bate: conteúdo é bônus de confirmação
                if sob_no_texto: score += 5
                if ano_no_texto: score += 3
                score += hits_cap * 2
                score += hits_livro
            else:
                # Sem match no filename: exige evidência mais forte no conteúdo
                if eh_compendio:
                    # Compêndio: aceita se o livro OU o autor do capítulo está no texto
                    if hits_livro >= 3 or (sob_no_texto and hits_cap >= 2):
                        if sob_no_texto: score += 8
                        if ano_no_texto: score += 3
                        score += hits_cap * 2
                        score += hits_livro * 2
                else:
                    # Obra simples: precisa de título específico no texto
                    if hits_cap >= 4:
                        if sob_no_texto: score += 8
                        if ano_no_texto: score += 3
                        score += hits_cap * 2

        if score > 0:
            candidatos.append((score, nome, texto))

    if not candidatos:
        if log_fn:
            log_fn(f"         → nenhum arquivo encontrado para {sobrenome} ({ano})")
        return "", ""

    candidatos.sort(key=lambda x: -x[0])
    melhor_score, nome_m, texto_m = candidatos[0]
    if log_fn:
        tipo = "compêndio" if eh_compendio else "obra"
        log_fn(f"         → {tipo} localizada: {nome_m[:65]} [score={melhor_score}]")
    return nome_m, texto_m


# ═══════════════════════════════════════════════════════════════════════════
# ENVIO DE ARQUIVOS NATIVOS AO CLAUDE (PDF como documento, DOCX como texto)
# ═══════════════════════════════════════════════════════════════════════════

_MAX_PDF_BYTES       = 28 * 1_000_000   # 28 MB por bloco (limite API ~32 MB)
_MAX_PAGES_PER_BLOCO = 50               # páginas por bloco de documento
_MAX_BLOCOS          = 8                # máximo de blocos por chamada


def _bloco_pdf(pdf_bytes: bytes, titulo: str) -> dict:
    return {
        "type": "document",
        "source": {
            "type": "base64",
            "media_type": "application/pdf",
            "data": base64.standard_b64encode(pdf_bytes).decode("ascii"),
        },
        "title": titulo,
    }


def _pdf_para_blocos(caminho: Path, contexto: str = "") -> list:
    """
    Converte um PDF em blocos de documento para a API do Claude.
    - PDF pequeno (≤ MAX_BYTES e ≤ MAX_PAGES): um único bloco com tudo.
    - PDF grande: seleciona as páginas mais relevantes ao contexto,
      divide em blocos de MAX_PAGES_PER_BLOCO e envia tudo que couber.
    Retorna lista de content-blocks prontos para incluir na mensagem.
    """
    try:
        with open(str(caminho), "rb") as f:
            raw = f.read()
    except Exception:
        return []

    try:
        import fitz  # PyMuPDF
    except ImportError:
        # Sem PyMuPDF: envia inteiro se couber, senão retorna vazio
        if len(raw) <= _MAX_PDF_BYTES:
            return [_bloco_pdf(raw, caminho.name)]
        return []

    try:
        doc = fitz.open(stream=raw, filetype="pdf")
        n_pages = len(doc)

        # PDF pequeno o suficiente para enviar inteiro
        if len(raw) <= _MAX_PDF_BYTES and n_pages <= _MAX_PAGES_PER_BLOCO * _MAX_BLOCOS:
            doc.close()
            return [_bloco_pdf(raw, caminho.name)]

        # PDF grande — seleciona páginas mais relevantes
        if contexto:
            palavras = {p for p in re.findall(r"[a-záéíóúâêîôûãõàç]{4,}", contexto.lower())
                        if p not in _STOPWORDS}
            scored = sorted(range(n_pages),
                            key=lambda i: -sum(1 for w in palavras
                                               if w in doc[i].get_text().lower()))
            selecionadas: set[int] = set()
            for pag in scored[:_MAX_PAGES_PER_BLOCO * _MAX_BLOCOS // 2]:
                for d in range(-3, 4):
                    p = pag + d
                    if 0 <= p < n_pages:
                        selecionadas.add(p)
            paginas = sorted(selecionadas)[:_MAX_PAGES_PER_BLOCO * _MAX_BLOCOS]
        else:
            paginas = list(range(min(n_pages, _MAX_PAGES_PER_BLOCO * _MAX_BLOCOS)))

        # Monta sub-PDFs por bloco de páginas
        blocos = []
        for start in range(0, len(paginas), _MAX_PAGES_PER_BLOCO):
            chunk = paginas[start:start + _MAX_PAGES_PER_BLOCO]
            sub = fitz.open()
            for p in chunk:
                sub.insert_pdf(doc, from_page=p, to_page=p)
            pdf_chunk = sub.tobytes()
            sub.close()
            titulo = (f"{caminho.name} "
                      f"(págs {chunk[0]+1}–{chunk[-1]+1} de {n_pages})")
            blocos.append(_bloco_pdf(pdf_chunk, titulo))

        doc.close()
        return blocos

    except Exception:
        # Qualquer falha: tenta enviar o raw se couber
        if len(raw) <= _MAX_PDF_BYTES:
            return [_bloco_pdf(raw, caminho.name)]
        return []


_REGRAS_ABNT = """
NORMAS ABNT PARA CITAÇÕES (NBR 10520:2023):
- Citação parentética: (SILVA, 2020) — sobrenome em MAIÚSCULAS
- Citação narrativa: Segundo Silva (2020) — só 1ª letra maiúscula
- 2 autores parentética: (SILVA; SOUZA, 2020) — separados por PONTO E VÍRGULA
- 4+ autores: (SILVA et al., 2020) — "et al." com ponto
- Citação direta: requer número de página — (SILVA, 2020, p. 45)
- Compêndio/coletânea: autoria é do AUTOR DO CAPÍTULO, não do organizador
- Apud: (AUTOR ORIGINAL apud AUTOR LIDO, ano) — use com moderação
"""

_PROMPT_CLAUDE = """Você é um verificador de integridade acadêmica especializado em normas ABNT.

{regras_abnt}

TRECHO DA DISSERTAÇÃO com a citação [{citacao}]:
\"\"\"{contexto}\"\"\"

REFERÊNCIA BIBLIOGRÁFICA COMPLETA: {referencia}
ARQUIVO FONTE: {arquivo} — leitura: {modo_leitura}
{info_compendio}
O conteúdo completo da obra está anexado acima como documento(s).

Analise com base no documento anexado:
1. O argumento da dissertação é coerente com o que a obra realmente afirma?
2. Há distorção de sentido, generalização indevida ou uso fora de contexto?
3. A ideia atribuída ao autor consta na obra? Em qual trecho específico?
4. Se for capítulo de compêndio: localize o capítulo do autor citado dentro do livro e verifique se o argumento corresponde ao que aquele capítulo afirma.

Responda SOMENTE com JSON (sem markdown):
{{"veredicto": "CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE", "justificativa": "1-2 frases explicando", "trecho_fonte": "trecho literal da obra até 200 chars"}}
"""

_PROMPT_CLAUDE_TEXTO = """Você é um verificador de integridade acadêmica especializado em normas ABNT.

{regras_abnt}

TRECHO DA DISSERTAÇÃO com a citação [{citacao}]:
\"\"\"{contexto}\"\"\"

REFERÊNCIA BIBLIOGRÁFICA COMPLETA: {referencia}
ARQUIVO FONTE: {arquivo} — leitura: {modo_leitura}
{info_compendio}

CONTEÚDO DA OBRA:
\"\"\"{fonte}\"\"\"

Analise com base no conteúdo acima:
1. O argumento da dissertação é coerente com o que a obra realmente afirma?
2. Há distorção de sentido, generalização indevida ou uso fora de contexto?
3. A ideia atribuída ao autor consta na obra? Em qual trecho específico?
4. Se for capítulo de compêndio: localize o capítulo do autor citado e verifique se o argumento corresponde ao que aquele capítulo afirma.

Responda SOMENTE com JSON (sem markdown):
{{"veredicto": "CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE", "justificativa": "1-2 frases explicando", "trecho_fonte": "trecho literal da obra até 200 chars"}}
"""

_LIMITE_TEXTO_COMPLETO = 150_000   # chars (~37 500 tokens) para fallback em texto


def verificar_claude(cit: Citacao, ref: Referencia, arquivo: str,
                     arq_path: Optional[Path], texto_fonte: str,
                     api_key: str) -> dict:
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        ext = arq_path.suffix.lower() if arq_path else ""
        blocos_doc: list = []
        modo = ""

        # ── Tenta enviar o arquivo nativo (PDF direto) ──────────────────
        if ext == ".pdf" and arq_path and arq_path.exists():
            blocos_doc = _pdf_para_blocos(arq_path, cit.contexto)
            if blocos_doc:
                n_pag_info = f"{len(blocos_doc)} bloco(s)"
                modo = f"PDF nativo — {n_pag_info}"

        # ── Fallback: texto extraído ─────────────────────────────────────
        if not blocos_doc:
            if not texto_fonte:
                return {"veredicto": "SEM_FONTE",
                        "justificativa": (
                            "PDF disponível mas não foi possível enviá-lo à API. "
                            "Instale PyMuPDF ('pip install PyMuPDF') para resolver."
                        ),
                        "trecho_fonte": ""}
            if len(texto_fonte) <= _LIMITE_TEXTO_COMPLETO:
                fonte_txt = texto_fonte
                modo = "texto completo"
            else:
                fonte_txt = _trecho_relevante(cit.contexto, texto_fonte,
                                              janela=_LIMITE_TEXTO_COMPLETO)
                modo = f"texto — seleção ({len(texto_fonte)//1024} KB total)"

            # Detecta compêndio para instruir Claude
            eh_comp = bool(re.search(r"\bIn:\s*", ref.texto, re.IGNORECASE))
            info_comp = (
                f"ATENÇÃO — COMPÊNDIO: Este é um capítulo de livro organizado. "
                f"Localize no arquivo o capítulo escrito por {ref.sobrenome} "
                f"e verifique o argumento nesse capítulo específico, não no texto do organizador."
                if eh_comp else ""
            )
            prompt_txt = _PROMPT_CLAUDE_TEXTO.format(
                regras_abnt=_REGRAS_ABNT,
                citacao=cit.texto,
                contexto=cit.contexto[:1200],
                referencia=ref.texto[:400],
                arquivo=arquivo,
                modo_leitura=modo,
                info_compendio=info_comp,
                fonte=fonte_txt,
            )
            content = [{"type": "text", "text": prompt_txt}]
        else:
            # PDF nativo: documentos + prompt de análise
            eh_comp = bool(re.search(r"\bIn:\s*", ref.texto, re.IGNORECASE))
            info_comp = (
                f"ATENÇÃO — COMPÊNDIO: Este é um capítulo de livro organizado. "
                f"Localize no arquivo o capítulo escrito por {ref.sobrenome} "
                f"e verifique o argumento nesse capítulo específico, não no texto do organizador."
                if eh_comp else ""
            )
            prompt_txt = _PROMPT_CLAUDE.format(
                regras_abnt=_REGRAS_ABNT,
                citacao=cit.texto,
                contexto=cit.contexto[:1200],
                referencia=ref.texto[:400],
                arquivo=arquivo,
                modo_leitura=modo,
                info_compendio=info_comp,
            )
            content = blocos_doc + [{"type": "text", "text": prompt_txt}]

        # Retry com backoff exponencial para rate limit (429)
        esperas = [30, 60, 120]
        for tentativa, espera in enumerate(esperas + [None], 1):
            try:
                resp = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1024,
                    messages=[{"role": "user", "content": content}],
                )
                break  # sucesso
            except Exception as exc:
                msg = str(exc)
                eh_rate_limit = "429" in msg or "rate_limit" in msg.lower() or "overloaded" in msg.lower()
                if eh_rate_limit and espera is not None:
                    time.sleep(espera)
                    continue
                raise  # outros erros: re-levanta imediatamente

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
        msg = str(e)[:200]
        if "429" in msg or "rate_limit" in msg.lower():
            msg = f"Rate limit da API atingido após 3 tentativas. Aguarde alguns minutos e tente novamente. ({msg})"
        return {"veredicto": "ERRO", "justificativa": msg, "trecho_fonte": ""}


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
<h1>📚 Verificador de Citações Acadêmicas</h1>
<p>Gerado em <b>{e(ts)}</b> &nbsp;|&nbsp; Arquivo: <em>{e(diss_path)}</em></p>

<div class="resumo">
  <div class="card" style="border-color:#3498db"><b>{len(resultado["citacoes"])}</b>Citações<br>encontradas</div>
  <div class="card" style="border-color:#8e44ad"><b>{len(resultado["referencias"])}</b>Referências<br>na lista</div>
  <div class="card" style="border-color:#27ae60"><b>{len(pareamentos)}</b>Pares<br>citação↔ref</div>
  <div class="card" style="border-color:#e74c3c"><b>{len(citadas_sr)}</b>Cit. sem<br>referência</div>
  <div class="card" style="border-color:#e67e22"><b>{len(refs_sc)}</b>Ref. sem<br>citação</div>
  <div class="card" style="border-color:#27ae60"><b>{len(corretas)}</b>Verificadas<br>OK</div>
  <div class="card" style="border-color:#f39c12"><b>{len(parciais)}</b>Parciais</div>
  <div class="card" style="border-color:#e74c3c"><b>{len(incorretas)}</b>Incorretas</div>
</div>
"""

    html += secao("⚠ Citações SEM referência na lista bibliográfica", citadas_sr, "#e74c3c",
        lambda c: (f'<b class="err">{e(c.texto)}</b>'
                   f'<details><summary><small>ver contexto</small></summary>'
                   f'<small>{e(c.contexto[:350])}</small></details>'))

    html += secao("⚠ Referências listadas SEM citação no texto", refs_sc, "#e67e22",
        lambda r: f'<span class="warn">{e(r.texto[:250])}</span>')

    if incorretas:
        html += secao("✗ INCORRETAS — argumento diverge da obra", incorretas, "#c0392b", card)
    if parciais:
        html += secao("~ PARCIAIS — citação com ressalvas", parciais, "#d35400", card)
    if corretas:
        html += secao("✓ CORRETAS — confirmadas pela obra", corretas, "#27ae60", card)
    if sem_fonte:
        html += secao("? SEM FONTE — arquivo não encontrado na pasta", sem_fonte, "#7f8c8d", card)

    if not verifics and pareamentos:
        html += f'<h2 style="color:#3498db">✓ Citações pareadas com referências ({len(pareamentos)})</h2>'
        html += '<p><em>Sem verificação semântica — informe a chave de API Claude para verificar a coerência dos argumentos.</em></p><ul>'
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

def analisar(diss_path: str, refs_dir: str, api_key: str, log_fn,
             stop_fn=None, sem_verificacao: bool = False) -> Optional[dict]:
    def L(msg):
        log_fn(msg)

    L("=" * 65)
    L("VERSÃO: 2025-06-02-v8.3")
    L("ETAPA 1 — Lendo a dissertação")
    L("=" * 65)
    texto = ler_docx(Path(diss_path))
    L(f"✓ {len(texto):,} caracteres extraídos")
    if len(texto) < 300:
        L("⚠ ERRO: texto muito curto! O arquivo pode ser uma imagem escaneada.")
        L(f"  Conteúdo encontrado: {repr(texto[:200])}")
        return None

    L("  Primeiras linhas:")
    for linha in texto.split("\n")[:5]:
        if linha.strip():
            L(f"  | {linha[:100]}")

    L("")
    L("  Identificando citações e referências na dissertação…")
    citacoes, idx_refs = encontrar_citacoes(texto, log_fn=L)
    L(f"  ✓ {len(citacoes)} citação(ões) encontrada(s)")
    if idx_refs >= 0:
        L(f"  ✓ Seção REFERÊNCIAS encontrada (linha {idx_refs})")
    else:
        L("  ⚠ Seção REFERÊNCIAS não encontrada no arquivo!")

    referencias = extrair_referencias(texto, idx_refs, log_fn=L)
    L(f"  ✓ {len(referencias)} entrada(s) na lista de referências")
    if referencias:
        L("  Primeiras 10 referências extraídas (sobrenomes | ano | início):")
        for r in referencias[:10]:
            L(f"    sob={r.sobrenomes} | {r.ano} | {r.texto[:80]}")

    L("")
    L("=" * 65)
    L("ETAPA 2 — Lendo arquivos de apoio (um a um)")
    L("=" * 65)
    extensoes = {".pdf", ".docx", ".doc", ".txt"}
    arquivos = sorted(
        [f for f in Path(refs_dir).rglob("*") if f.is_file() and f.suffix.lower() in extensoes],
        key=lambda f: f.name.lower()
    )
    L(f"  {len(arquivos)} arquivo(s) encontrado(s) na pasta")
    L("")

    material       = {}   # nome → texto extraído  (para matching por conteúdo)
    material_paths = {}   # nome → Path            (para envio nativo ao Claude)
    pdf_sem_texto  = 0
    for i, arq in enumerate(arquivos, 1):
        if stop_fn and stop_fn():
            L("\n[INTERROMPIDO]")
            return None
        t = ler_arquivo(arq)
        is_pdf = arq.suffix.lower() == ".pdf"
        ok_texto = bool(t) and len(t) > 50 and not t.startswith("[")

        if ok_texto:
            material[arq.name]       = t
            material_paths[arq.name] = arq
            tipo = "PDF" if is_pdf else arq.suffix.upper().lstrip(".")
            L(f"  [{i:>3}/{len(arquivos)}] ✓ {arq.name}  ({len(t):,} chars, {tipo})")
        elif is_pdf and arq.exists() and arq.stat().st_size > 0:
            # PDF sem extração de texto (PyMuPDF ausente), mas válido para envio nativo à API
            material[arq.name]       = ""
            material_paths[arq.name] = arq
            pdf_sem_texto += 1
            L(f"  [{i:>3}/{len(arquivos)}] ~ {arq.name}  (PDF disponível — envio nativo à API)")
        else:
            L(f"  [{i:>3}/{len(arquivos)}] ✗ {arq.name}  — falha na leitura")

    com_texto = len(material) - pdf_sem_texto
    L(f"\n  ✓ {len(material)} arquivo(s) disponível(is): {com_texto} com texto, {pdf_sem_texto} somente PDF nativo")
    if pdf_sem_texto > 0:
        L("  ℹ Para extrair texto de PDFs (melhora o matching), instale:")
        L("    pip install PyMuPDF")
        L("  Os PDFs serão enviados nativamente ao Claude mesmo sem texto extraído.")

    L("")
    L("=" * 65)
    L("ETAPA 3 — Cruzando citações com referências")
    L("=" * 65)
    cruzamento = cruzar(citacoes, referencias, log_fn=L)
    pareamentos      = cruzamento["pareamentos"]
    citadas_sem_ref  = cruzamento["citadas_sem_ref"]
    refs_sem_citacao = cruzamento["refs_sem_citacao"]
    L(f"  ✓ Pares encontrados:             {len(pareamentos)}")
    L(f"  ✓ Citadas SEM referência:        {len(citadas_sem_ref)}")
    L(f"  ✓ Referências SEM citação:       {len(refs_sem_citacao)}")

    if citadas_sem_ref:
        L("\n  Citações sem referência:")
        for c in citadas_sem_ref[:10]:
            L(f"    → {c.texto}")
        if len(citadas_sem_ref) > 10:
            L(f"    … e mais {len(citadas_sem_ref) - 10}")

    verificacoes = []
    L("")
    L("=" * 65)
    L("ETAPA 4 — Verificação semântica com IA")
    L("=" * 65)
    L(f"  Chave API informada:      {'SIM' if api_key.strip() else 'NÃO — sem verificação semântica'}")
    L(f"  Checkbox 'só cruzamento': {'SIM — semântica desativada' if sem_verificacao else 'não'}")
    L(f"  Arquivos de apoio lidos:  {len(material)}")
    L(f"  Pares citação↔referência: {len(pareamentos)}")

    if sem_verificacao:
        L("  [Verificação semântica desativada pelo checkbox]")
    elif not api_key.strip():
        L("  [Sem chave API — informe a chave Anthropic no campo correspondente]")
    elif not material:
        L("")
        L("  ⚠ ATENÇÃO: Chave API fornecida mas NENHUM arquivo de apoio lido!")
        L("  → Coloque os PDFs/DOCXs das obras na pasta de referências.")
    else:
        # ── Verifica citações que TÊM referência pareada ────────────────
        candidatos_semantica = []  # (cit, ref_ou_None, arq, arq_path, txt)

        if pareamentos:
            L(f"  → Localizando arquivos para {len(pareamentos)} par(es) citação↔referência...")
        for cit, ref in pareamentos:
            # Usa o texto COMPLETO da referência para matching por título
            arq, txt = _buscar_fonte(cit.autores[0], cit.ano, ref.texto, material, log_fn=L)
            arq_path = material_paths.get(arq) if arq else None
            candidatos_semantica.append((cit, ref, arq, arq_path, txt))

        # ── Também verifica citadas_sem_ref quando há arquivo disponível ─
        sem_ref_com_fonte = []
        for cit in citadas_sem_ref:
            arq, txt = _buscar_fonte(cit.autores[0], cit.ano,
                                     " ".join(cit.autores), material, log_fn=L)
            if arq:
                arq_path = material_paths.get(arq) if arq else None
                sem_ref_com_fonte.append((cit, None, arq, arq_path, txt))

        if sem_ref_com_fonte:
            L(f"  → {len(sem_ref_com_fonte)} citação(ões) sem referência têm arquivo na pasta — também serão verificadas.")
            candidatos_semantica.extend(sem_ref_com_fonte)

        total_verif = len(candidatos_semantica)
        if total_verif == 0:
            L("")
            L("  ⚠ Nenhum arquivo de apoio correspondeu às citações/referências.")
            L("  → Verifique se os nomes dos arquivos contêm o sobrenome do autor.")
        else:
            L(f"  → Iniciando verificação semântica de {total_verif} citação(ões)...")
            L("")
            for i, (cit, ref, arq, arq_path, txt) in enumerate(candidatos_semantica, 1):
                if stop_fn and stop_fn():
                    L("\n[INTERROMPIDO]")
                    break
                tam = f"{arq_path.stat().st_size // 1024} KB" if arq_path else "—"
                L(f"  [{i:>3}/{total_verif}] {cit.texto[:70]}")
                # Verifica se o arquivo foi realmente encontrado (arq pode estar vazio
                # mas txt pode ser "" em PDFs sem extração — o que ainda é válido para API)
                if not arq:
                    ref_sob = ref.sobrenome if ref else cit.autores[0]
                    ref_ano = ref.ano if ref else cit.ano
                    L(f"         → arquivo não encontrado para {ref_sob} ({ref_ano})")
                    verificacoes.append({
                        "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                        "contexto": cit.contexto,
                        "referencia": ref.texto if ref else f"{cit.autores[0]} ({cit.ano})",
                        "arquivo_fonte": "", "veredicto": "SEM_FONTE",
                        "justificativa": f"Arquivo de {ref_sob} ({ref_ano}) não encontrado na pasta",
                        "trecho_fonte": "",
                    })
                    continue
                modo_txt = f"texto {len(txt):,} chars" if txt else "somente PDF nativo"
                L(f"         → fonte: {arq}  ({tam}, {modo_txt})")
                # Monta ref_obj para verificar_claude (usa dados reais ou proxy da citação)
                if ref is None:
                    ref_obj = Referencia(
                        texto=f"{cit.autores[0]} ({cit.ano}) — sem entrada na lista de referências",
                        sobrenome=cit.autores[0],
                        sobrenomes=cit.autores,
                        ano=cit.ano,
                        titulo=" ".join(cit.autores),
                    )
                else:
                    ref_obj = ref
                v = verificar_claude(cit, ref_obj, arq, arq_path, txt, api_key)
                emoji = {"CORRETO": "✓", "INCORRETO": "✗", "PARCIAL": "⚠",
                         "SEM_FONTE": "?", "ERRO": "!"}.get(v.get("veredicto", ""), "?")
                L(f"         {emoji} {v.get('veredicto')} — {v.get('justificativa','')[:80]}")
                verificacoes.append({
                    "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                    "contexto": cit.contexto,
                    "referencia": ref_obj.texto,
                    "arquivo_fonte": arq, **v,
                })
                # Pausa entre chamadas para evitar rate limit (3s mínimo)
                if i < total_verif:
                    time.sleep(3)

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
    janela.title("Verificador de Citações Acadêmicas")
    janela.geometry("820x680")
    janela.resizable(True, True)
    try:
        janela.iconbitmap(default="")
    except Exception:
        pass

    # ── Painel superior ──────────────────────────────────────────────────
    topo = tk.Frame(janela, bg="#1a2744", padx=12, pady=10)
    topo.pack(fill="x")
    tk.Label(topo, text="📚  Verificador de Citações Acadêmicas",
             bg="#1a2744", fg="white", font=("Arial", 15, "bold")).pack(anchor="w")
    tk.Label(topo, text="Lê a dissertação · Lê os arquivos de apoio · Cruza e verifica",
             bg="#1a2744", fg="#aac", font=("Arial", 10)).pack(anchor="w")

    # ── Formulário ───────────────────────────────────────────────────────
    form = tk.Frame(janela, padx=16, pady=10)
    form.pack(fill="x")

    var_diss  = tk.StringVar()
    var_refs  = tk.StringVar()
    var_api   = tk.StringVar()
    var_sem   = tk.BooleanVar(value=False)

    def _campo(parent, texto, var, comando, row):
        tk.Label(parent, text=texto, anchor="w", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=4)
        tk.Entry(parent, textvariable=var, width=58, font=("Arial", 10)).grid(
            row=row, column=1, padx=(6, 4), pady=4, sticky="ew")
        tk.Button(parent, text="Procurar…", command=comando,
                  font=("Arial", 9)).grid(row=row, column=2, pady=4)

    def escolher_diss():
        p = filedialog.askopenfilename(
            title="Selecione a dissertação",
            filetypes=[("Word", "*.docx *.doc"), ("Todos", "*.*")])
        if p:
            var_diss.set(p)

    def escolher_refs():
        p = filedialog.askdirectory(title="Selecione a pasta com os arquivos de referência")
        if p:
            var_refs.set(p)

    form.columnconfigure(1, weight=1)
    _campo(form, "Dissertação (.docx):", var_diss, escolher_diss, 0)
    _campo(form, "Pasta de referências:", var_refs, escolher_refs, 1)

    tk.Label(form, text="Chave API Claude\n(opcional):", anchor="w",
             font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=4)
    tk.Entry(form, textvariable=var_api, width=58, font=("Arial", 10),
             show="•").grid(row=2, column=1, padx=(6, 4), pady=4, sticky="ew")
    tk.Button(form, text="Ver/Ocultar",
              command=lambda: e_api.config(show="" if e_api.cget("show") == "•" else "•"),
              font=("Arial", 9)).grid(row=2, column=2, pady=4)

    # Substituir a entry genérica pela que tem referência
    e_api = form.grid_slaves(row=2, column=1)[0]

    tk.Checkbutton(form, text="Apenas cruzamento (sem verificação IA, mais rápido)",
                   variable=var_sem, font=("Arial", 10)).grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

    # ── Botões de ação ───────────────────────────────────────────────────
    btn_frame = tk.Frame(janela, padx=16)
    btn_frame.pack(fill="x")

    btn_iniciar = tk.Button(btn_frame, text="▶  Iniciar análise",
                            font=("Arial", 11, "bold"), bg="#27ae60", fg="white",
                            padx=14, pady=6)
    btn_parar   = tk.Button(btn_frame, text="⏹  Parar",
                            font=("Arial", 11), bg="#e74c3c", fg="white",
                            padx=14, pady=6, state="disabled")
    btn_relat   = tk.Button(btn_frame, text="🌐  Abrir relatório HTML",
                            font=("Arial", 11), bg="#2980b9", fg="white",
                            padx=14, pady=6, state="disabled")

    btn_iniciar.pack(side="left", padx=(0, 8), pady=8)
    btn_parar.pack(side="left", padx=(0, 8), pady=8)
    btn_relat.pack(side="left", pady=8)

    # ── Barra de progresso ───────────────────────────────────────────────
    prog_var = tk.DoubleVar()
    prog_bar = ttk.Progressbar(janela, variable=prog_var, mode="indeterminate")
    prog_bar.pack(fill="x", padx=16, pady=(0, 4))

    # ── Log ──────────────────────────────────────────────────────────────
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
            messagebox.showerror("Erro", "Selecione um arquivo de dissertação válido.")
            return
        if not refs or not Path(refs).is_dir():
            messagebox.showerror("Erro", "Selecione uma pasta de referências válida.")
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
                    sem_verificacao=var_sem.get(),
                )
                if resultado:
                    html_str = gerar_html(resultado)
                    saida = Path(diss).parent / "relatorio_citacoes.html"
                    saida.write_text(html_str, encoding="utf-8")
                    _resultado["html_path"] = str(saida).replace("\\", "/")
                    log("")
                    log("=" * 65)
                    log("✓ ANÁLISE CONCLUÍDA!")
                    log(f"  Relatório salvo em: {saida}")
                    log("=" * 65)
                    def _ativar():
                        btn_relat.configure(state="normal")
                        prog_bar.stop()
                        prog_var.set(100)
                        messagebox.showinfo("Concluído",
                            f"Análise finalizada!\n\n"
                            f"• {len(resultado['citacoes'])} citações encontradas\n"
                            f"• {len(resultado['referencias'])} referências na lista\n"
                            f"• {len(resultado['pareamentos'])} pares citação↔referência\n"
                            f"• {len(resultado['citadas_sem_ref'])} citadas sem referência\n"
                            f"• {len(resultado['refs_sem_citacao'])} referências sem citação\n\n"
                            f"Relatório salvo em:\n{saida}")
                    janela.after(0, _ativar)
                else:
                    log("\n[Análise não concluída — verifique os erros acima]")
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
        log("\n[Solicitando parada…]")

    btn_iniciar.configure(command=iniciar)
    btn_parar.configure(command=parar)

    log("Bem-vindo ao Verificador de Citações Acadêmicas  [versão 2025-06-02-v8.3]")
    log("1. Selecione o arquivo da dissertação (.docx)")
    log("2. Selecione a pasta com os arquivos de referência (PDF, DOCX, TXT)")
    log("3. Informe a chave API Claude (opcional) para verificação de coerência")
    log("4. Clique em 'Iniciar análise'")
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
    print(f"\nRelatório salvo: {saida}")
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
            print("tkinter não disponível. Use: python analisar.py dissertacao.docx pasta_refs/")
            sys.exit(1)
