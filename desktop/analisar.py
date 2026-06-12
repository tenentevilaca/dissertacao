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


def ler_pdf_paginas(caminho: Path) -> list:
    """Retorna o texto de cada página separadamente — [texto_pag_1, texto_pag_2, ...].

    Diferente de ler_pdf(), preserva a fronteira de cada página para permitir
    apontar o número exato onde um trecho aparece (PDF tem paginação fixa
    gravada no arquivo; .docx não — por isso essa função existe só para PDF).
    """
    try:
        import fitz
        doc = fitz.open(str(caminho))
        return [page.get_text() for page in doc]
    except Exception:
        return []


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


_RE_FRONTEIRA_FRASE = re.compile(r'[.!?](?=\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ"“]|\s*$)')


def _trecho_em_torno(contexto: str, alvo: str, janela: int = 350) -> str:
    """Recorta um trecho de `contexto` delimitado pelas FRONTEIRAS DE FRASE
    em torno da ocorrência da própria citação (`alvo`), em vez de um corte
    fixo de caracteres.

    Dois problemas resolvidos:
    1) Como `_contexto` inclui ~300 caracteres ANTES do parágrafo da citação,
       um corte fixo no INÍCIO do contexto frequentemente mostrava só o
       parágrafo ANTERIOR — fazendo parecer que a citação "pertencia" a
       outro trecho/argumento, quando na verdade só aparecia mais adiante.
    2) Um corte fixo no FIM podia avançar para a frase/parágrafo SEGUINTE,
       que discute outro raciocínio — apresentando como "contexto da
       citação" um trecho sem nenhuma relação com o argumento que ela
       sustenta.

    Por isso o lado posterior NUNCA ultrapassa o fim da própria frase que
    contém a citação (onde o argumento citado termina); o lado anterior
    pode incluir frases-gancho prévias (até `janela` caracteres), mas
    sempre encaixado em fronteiras de frase — nunca cortando uma oração
    ao meio.
    """
    pos = contexto.find(alvo)
    if pos < 0:
        return contexto[:janela]
    fim_alvo = pos + len(alvo)

    # Lado posterior: vai só até o fim da frase que contém a citação —
    # texto além disso pertence a outro raciocínio.
    janela_pos = contexto[fim_alvo:fim_alvo + janela]
    m_fim = _RE_FRONTEIRA_FRASE.search(janela_pos)
    fim = fim_alvo + (m_fim.end() if m_fim else len(janela_pos))

    # Lado anterior: até `janela` caracteres, encaixado na fronteira de
    # frase mais próxima para preservar o "gancho" do argumento sem
    # cortar uma oração ao meio.
    inicio_bruto = max(0, pos - janela)
    fronteiras_antes = list(_RE_FRONTEIRA_FRASE.finditer(contexto, inicio_bruto, pos))
    inicio = fronteiras_antes[-1].end() if fronteiras_antes else inicio_bruto

    trecho = contexto[inicio:fim].strip()
    if inicio > 0:
        trecho = "[...] " + trecho
    if fim < len(contexto):
        trecho = trecho + " [...]"
    return trecho



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
    # 4ª alt: LAS FUERZAS: periódico... (entrada com título em caixa-alta seguido
    # de dois-pontos — convenção ABNT para obras sem autoria pessoal/institucional,
    # onde a entrada é alfabetada pelo próprio título). Limites de tamanho mínimo
    # (1ª palavra ≥3 maiúsculas, demais ≥3 ao todo) evitam falso-positivo em
    # fragmentos comuns como "DISPONÍVEL EM:" (EM tem 2 letras, não casa).
    # 5ª alt: OSCE – Organization for Security... (entrada institucional iniciada
    # pela própria sigla, seguida de travessão e do nome por extenso — outra
    # convenção ABNT comum em obras de organismos internacionais).
    _RE_NOVA = re.compile(
        r"^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+[,;]"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,}[a-záéíóúâêîôûãõàç\'\-]*\.\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,}(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç]*)+\s*(?:\([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,12}\)\s*)?[,.]"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{3,}(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç]{2,})+\s*:"
        r"|^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,12}\s*[–—\-]\s*[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][a-záéíóúâêîôûãõàç]"
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
            # Nome institucional grafado por extenso, sem sigla declarada entre
            # parênteses: gera a sigla pelas iniciais das palavras significativas
            # para casar com citações abreviadas — ex.: "Fórum Brasileiro de
            # Segurança Pública" → "FBSP"; "Organização para a Segurança e
            # Cooperação na Europa" → "OSCE".
            # Referências institucionais ABNT costumam aninhar o órgão dentro da
            # jurisdição — ex.: "UNITED STATES. Department of Justice; Department
            # of Homeland Security. Baseline capabilities..." É o ÓRGÃO (DOJ, DHS),
            # não a jurisdição (US), que aparece na citação no texto. Por isso
            # percorre cada segmento (separado por "." ou ";") gerando uma sigla
            # para cada um que tenha "cara" de nome institucional (todas as
            # palavras iniciam maiúsculas ou são stopwords), parando no primeiro
            # segmento que pareça título (contém palavra minúscula não-stopword,
            # ex.: "Baseline capabilities for state...").
            _stop_sigla = {'de', 'da', 'do', 'das', 'dos', 'e', 'a', 'o', 'as',
                           'os', 'para', 'em', 'no', 'na', 'nos', 'nas',
                           'com', 'por', 'à', 'ao', 'um', 'uma',
                           'of', 'the', 'and', 'for', 'in', 'on', 'to'}
            for seg in re.split(r'[\.;]\s+', t[:300])[:5]:
                palavras = seg.split()
                if not palavras or len(palavras) > 8:
                    break
                if not all(re.match(r'^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]', p) or p.lower() in _stop_sigla
                           for p in palavras):
                    break
                # Gera duas variantes por segmento: excluindo e incluindo "of" —
                # convenção inglesa é inconsistente (Department OF Justice → DOJ
                # inclui o "of", mas Department OF Homeland Security → DHS não),
                # e não há como saber qual sem a sigla declarada na própria obra.
                for incluir_of in (False, True):
                    _stop = _stop_sigla - {'of'} if incluir_of else _stop_sigla
                    sigla_gerada = "".join(p[0].upper() for p in palavras
                                           if p.lower() not in _stop)
                    if 2 <= len(sigla_gerada) <= 10 and sigla_gerada not in resultado:
                        resultado.append(sigla_gerada)
        # Extrai siglas entre parênteses em QUALQUER ponto da referência — não só
        # no início — para cobrir entradas como "Organização para a Segurança e
        # Cooperação na Europa (OSCE)." onde a sigla vem depois do nome por extenso.
        # (Padrão restrito a parênteses propositalmente: buscar siglas soltas após
        # travessão/dois-pontos também capturava trechos de URLs/nomes de arquivo
        # nas referências on-line, ex.: "...uploads/2025/relatorio-BRASIL.pdf",
        # gerando pareamentos falsos.)
        siglas = re.findall(r'\(([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ]{2,12})\)', t)
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
    # Sufixos geracionais comuns em sobrenomes brasileiros (FILHO, JÚNIOR,
    # NETO, SOBRINHO) são curtos e genéricos — tolerância a distância de
    # edição os faz colidir com sobrenomes não relacionados por coincidência
    # (ex.: "Fialho" ↔ "Filho", distância de edição = 1, mas são pessoas
    # diferentes). Por isso só casam por igualdade exata.
    _sufixos_genericos = {'filho', 'junior', 'jr', 'neto', 'sobrinho'}
    if na in _sufixos_genericos or nb in _sufixos_genericos:
        return False
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
    # Observação: NÃO deduplicamos pareamentos por referência (id(ref)). Duas
    # citações distintas — ex.: "(FBSP, 2024)" e "(FBSP, 2025)" — podem casar
    # com a MESMA entrada bibliográfica (via fuzzy ±1 ano) e ambas precisam
    # aparecer no relatório e ser verificadas; suprimir a segunda fazia a
    # citação "desaparecer" silenciosamente (nem pareada, nem listada como
    # sem referência).

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
                pareamentos.append((c, ref))
                encontrou = True
                break

            # 2) Fuzzy: sobrenome similar + mesmo ano
            if not encontrou:
                for r in referencias:
                    if r.ano == c.ano and any(_fuzzy_sobrenome(sob, rs) for rs in r.sobrenomes):
                        for s in r.sobrenomes:
                            chaves_citadas.add(f"{s.upper()}_{r.ano}")
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
                        except (ValueError, TypeError, AttributeError):
                            continue
                        if (abs(ano_cit - ano_ref) <= 1
                                and any(_fuzzy_sobrenome(sob, rs) for rs in r.sobrenomes)):
                            for s in r.sobrenomes:
                                chaves_citadas.add(f"{s.upper()}_{r.ano}")
                            pareamentos.append((c, r))
                            encontrou = True
                            break
                except (ValueError, TypeError):
                    pass

            # 3b) Referência sem data declarada ("s.d." — convenção ABNT para
            #     obra sem ano identificável, ex.: "[S.l.: s.n.], [s.d.]"):
            #     a referência não pode "discordar" de um ano específico citado
            #     quando ela própria não declara nenhum — por isso casa pelo
            #     sobrenome (exato ou aproximado), independentemente do ano da
            #     citação. Evita falso-negativo "sem referência" em casos como
            #     "Fialho (2019)" citado no texto vs. referência "FIALHO...
            #     [s.d.]" na bibliografia (o ano só existe de um lado).
            if not encontrou:
                for r in referencias:
                    if r.ano == "s.d." and (
                        sob.upper() in (s.upper() for s in r.sobrenomes)
                        or any(_fuzzy_sobrenome(sob, rs) for rs in r.sobrenomes)
                    ):
                        for s in r.sobrenomes:
                            chaves_citadas.add(f"{s.upper()}_{r.ano}")
                        pareamentos.append((c, r))
                        encontrou = True
                        break

            # 4) Sigla institucional: o "sobrenome" citado é uma sigla curta
            #    (ex.: OSCE, ONU, FBSP, OEA) que pode não abrir a entrada da
            #    referência — a obra pode estar indexada pelo nome por extenso,
            #    com a sigla aparecendo entre parênteses, após travessão etc.
            #    Procuramos a sigla como palavra isolada em qualquer parte do
            #    texto da referência, com o mesmo ano (ou ano ±1).
            if not encontrou and 2 <= len(sob) <= 10 and sob.isalpha():
                sigla_re = re.compile(rf'\b{re.escape(sob)}\b')
                try:
                    ano_cit = int(c.ano[:4])
                except (ValueError, TypeError):
                    ano_cit = None
                for r in referencias:
                    # Ignora URLs/links de acesso: trechos como
                    # ".../uploads/2025/relatorio-BRASIL.pdf" podem conter
                    # substrings maiúsculas que colidem com siglas/sobrenomes
                    # e gerar pareamentos falsos.
                    texto_busca = re.sub(r'https?://\S+|www\.\S+', ' ', r.texto)
                    if not sigla_re.search(texto_busca):
                        continue
                    try:
                        ano_ref = int(r.ano[:4])
                    except (ValueError, TypeError, AttributeError):
                        ano_ref = None
                    ano_ok = (r.ano == c.ano or
                              (ano_cit is not None and ano_ref is not None
                               and abs(ano_cit - ano_ref) <= 1))
                    if ano_ok:
                        for s in r.sobrenomes:
                            chaves_citadas.add(f"{s.upper()}_{r.ano}")
                        pareamentos.append((c, r))
                        _log(f"    [SIGLA-MATCH] {c.texto} ↔ {r.texto[:80]}")
                        encontrou = True
                        break

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


def _sanitizar(texto: str) -> str:
    """Remove bytes nulos e caracteres de controle que invalidam JSON na API."""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', ' ', texto)


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

    # ── Document-frequency das palavras do título entre os candidatos ──
    # Palavras genéricas do domínio (ex.: "intelligence", "policing") tendem
    # a aparecer em quase todos os arquivos de apoio. Para evitar que esses
    # "hits" frágeis sustentem um match somente-por-conteúdo, calculamos a
    # fração de arquivos que contém cada palavra e descontamos as muito
    # frequentes (estilo IDF simplificado).
    todas_palavras = set(palavras_cap) | set(palavras_livro)
    n_arquivos = max(len(material), 1)
    df = {}
    for p in todas_palavras:
        cnt = 0
        for texto in material.values():
            if texto and p in texto.lower():
                cnt += 1
        df[p] = cnt / n_arquivos

    def _palavra_distintiva(p: str) -> bool:
        # Considera "distintiva" se aparece em menos de 60% dos arquivos
        return df.get(p, 0.0) < 0.6

    TAMANHO_CABECALHO = 2000  # janela do "início do documento" / título-página

    candidatos = []
    for nome, texto in material.items():
        score = 0
        nome_l = nome.lower()
        tl = texto.lower() if texto else ""
        tl_inicio = tl[:TAMANHO_CABECALHO]

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

            # Hits "distintivos" (descontando palavras genéricas de alta
            # frequência no corpus de apoio) e hits no cabeçalho/início do
            # documento (proxy de que o arquivo É a própria obra, não apenas
            # cita-a em uma bibliografia)
            hits_cap_dist = sum(1 for p in palavras_cap
                                 if p in tl and _palavra_distintiva(p))
            hits_livro_dist = sum(1 for p in palavras_livro
                                   if p in tl and _palavra_distintiva(p))
            hits_cap_inicio = sum(1 for p in palavras_cap if p in tl_inicio)
            hits_livro_inicio = sum(1 for p in palavras_livro if p in tl_inicio)

            if nome_no_arquivo:
                # Filename já bate: conteúdo é bônus de confirmação
                if sob_no_texto: score += 5
                if ano_no_texto: score += 3
                score += hits_cap * 2
                score += hits_livro
            else:
                # Sem match no filename: exige evidência mais forte no conteúdo
                if eh_compendio:
                    # Compêndio: aceita se o livro (com palavras distintivas
                    # e perto do início do doc) OU o autor+título do capítulo
                    # está no texto
                    aceita_livro = (hits_livro_dist >= 3 and hits_livro_inicio >= 2)
                    aceita_cap   = (sob_no_texto and hits_cap_dist >= 2 and hits_cap_inicio >= 1)
                    if aceita_livro or aceita_cap:
                        if sob_no_texto: score += 8
                        if ano_no_texto: score += 3
                        score += hits_cap * 2
                        score += hits_livro * 2
                else:
                    # Obra simples sem match no filename: o título é quem
                    # individualiza a obra. Sobrenome+ano sozinhos não bastam,
                    # pois qualquer artigo que cite o autor na bibliografia
                    # terá ambos no texto. Exige palavras DISTINTIVAS do
                    # título (não genéricas do domínio) e que apareçam perto
                    # do início do documento — indício de que o arquivo É a
                    # própria obra (título/capa), não apenas a cita em sua
                    # bibliografia.
                    if hits_cap_dist >= 3 and hits_cap_inicio >= 2:
                        score += hits_cap * 2
                        if sob_no_texto: score += 5
                        if ano_no_texto: score += 3
                    elif (hits_cap_dist >= 2 and hits_cap_inicio >= 2
                          and sob_no_texto and ano_no_texto):
                        # Evidência fraca de título, mas reforçada por
                        # autor+ano combinados no texto e presença no início
                        score += hits_cap * 2 + 4

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

ATENÇÃO — TRATE FORMATAÇÃO E SEMÂNTICA COMO COISAS COMPLETAMENTE SEPARADAS:
1. PROBLEMA DE FORMATAÇÃO (forma): a citação no texto não segue à risca o padrão
   ABNT acima — ex.: vírgula em vez de ponto e vírgula entre autores, parênteses
   faltando ou sobrando, espaçamento, "et al" sem ponto, ano colado ao sobrenome,
   maiúsculas/minúsculas trocadas etc. Isso é só uma questão de FORMA.
2. PROBLEMA SEMÂNTICO (conteúdo): o argumento atribuído ao autor não corresponde
   ao que a obra realmente afirma — distorção de sentido, generalização indevida,
   uso fora de contexto, inversão do argumento. Isso sim é DIVERGÊNCIA.
REGRA DE OURO: um desvio de formatação NUNCA torna uma citação "INCORRETA" ou
"PARCIAL" do ponto de vista do argumento. São defeitos de natureza diferente —
relate-os em campos separados, sem deixar que um contamine o outro.
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
{{"veredicto": "CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE", "justificativa": "1-2 frases sobre a fidelidade do argumento à obra", "trecho_fonte": "trecho literal da obra até 200 chars", "formato_abnt": {{"conforme": true|false, "observacao": "se não conforme, descreva o desvio em <=1 frase; se conforme, string vazia"}}}}
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
{{"veredicto": "CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE", "justificativa": "1-2 frases sobre a fidelidade do argumento à obra", "trecho_fonte": "trecho literal da obra até 200 chars", "formato_abnt": {{"conforme": true|false, "observacao": "se não conforme, descreva o desvio em <=1 frase; se conforme, string vazia"}}}}
"""

_LIMITE_TEXTO_SEMANTICA  = 30_000   # janela para 1 citação
_LIMITE_TEXTO_GRUPO      = 50_000   # janela maior quando múltiplas citações da mesma obra
_MAX_CITS_POR_CHAMADA    = 10       # máximo de citações por chamada à API


def verificar_grupo_claude(grupo: list, arquivo: str,
                            arq_path: Optional[Path], texto_fonte: str,
                            api_key: str, log_fn=None) -> list:
    """
    Verifica todas as citações de uma mesma obra em uma única chamada.
    `grupo` é uma lista de (Citacao, Referencia|None, arq, arq_path, txt).
    Retorna lista de dicts com veredicto na mesma ordem do grupo.
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)

    import anthropic
    client = anthropic.Anthropic(api_key=api_key, max_retries=5)

    todos_contextos = " ".join(item[0].contexto for item in grupo)

    # Monta lista de citações para o prompt
    def _montar_lista(sub):
        linhas = []
        for i, (cit, ref, _, _, _) in enumerate(sub, 1):
            ref_txt = ref.texto[:300] if ref else f"{cit.autores[0]} ({cit.ano})"
            comp = ""
            if ref and re.search(r"\bIn:\s*", ref.texto, re.IGNORECASE):
                comp = f" [COMPÊNDIO — capítulo de {ref.sobrenome}]"
            linhas.append(
                f"[{i}] CITAÇÃO: {cit.texto}\n"
                f"    CONTEXTO: {_trecho_em_torno(cit.contexto, cit.texto, 350)}\n"
                f"    REFERÊNCIA: {ref_txt}{comp}"
            )
        return "\n\n".join(linhas)

    instrucao_final = (
        "Para cada citação avalie DUAS coisas, de forma independente:\n"
        "  (a) SEMÂNTICA — o argumento da dissertação é coerente com o que a obra "
        "realmente afirma? (define o campo \"veredicto\")\n"
        "  (b) FORMATAÇÃO — a citação no texto segue o padrão ABNT descrito acima? "
        "(define o campo \"formato_abnt\", à parte)\n"
        "NÃO deixe desvios de formatação influenciarem o veredicto semântico — "
        "uma citação fora do padrão ABNT mas com argumento fiel à obra é CORRETA.\n"
        "Responda SOMENTE com array JSON (sem markdown), um objeto por citação, na mesma ordem:\n"
        '[{"idx":1,"veredicto":"CORRETO"|"INCORRETO"|"PARCIAL"|"SEM_FONTE",'
        '"justificativa":"1 frase sobre a fidelidade do argumento à obra",'
        '"trecho_fonte":"trecho literal ≤150 chars",'
        '"formato_abnt":{"conforme":true|false,'
        '"observacao":"se não conforme, descreva o desvio em ≤1 frase; se conforme, string vazia"}},...]'
    )

    def _chamar_api(sub, content):
        max_tok = min(4096, 400 * len(sub) + 512)
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=max_tok,
            messages=[{"role": "user", "content": content}],
        )
        raw = re.sub(r"^```(?:json)?", "", resp.content[0].text.strip()).strip("`").strip()
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            objs = re.findall(r'\{[^{}]+\}', raw, re.DOTALL)
            parsed = []
            for s in objs:
                try:
                    parsed.append(json.loads(s))
                except Exception:
                    pass
        if isinstance(parsed, dict):
            parsed = [parsed]
        parsed = sorted(parsed, key=lambda x: x.get("idx", 999)) if isinstance(parsed, list) else []
        while len(parsed) < len(sub):
            parsed.append({"veredicto": "ERRO", "justificativa": "Resposta incompleta da IA", "trecho_fonte": ""})
        for p in parsed:
            if not isinstance(p.get("formato_abnt"), dict):
                p["formato_abnt"] = {"conforme": True, "observacao": ""}
        return parsed[:len(sub)]

    def _erro_lista(sub, msg):
        return [{"veredicto": "ERRO", "justificativa": msg[:150], "trecho_fonte": "",
                 "formato_abnt": {"conforme": True, "observacao": ""}} for _ in sub]

    resultados = []
    # Divide em lotes de MAX_CITS_POR_CHAMADA
    for inicio in range(0, len(grupo), _MAX_CITS_POR_CHAMADA):
        sub = grupo[inicio:inicio + _MAX_CITS_POR_CHAMADA]
        sub_contextos = " ".join(item[0].contexto for item in sub)
        lista_str = _montar_lista(sub)

        try:
            if texto_fonte:
                fonte_txt = _sanitizar(_trecho_relevante(sub_contextos + " " + todos_contextos,
                                              texto_fonte, janela=_LIMITE_TEXTO_GRUPO))
                modo = f"texto {len(fonte_txt)//1024}KB/{len(texto_fonte)//1024}KB"
                _log(f"     → texto {len(fonte_txt):,} chars de {len(texto_fonte):,} — lote {inicio//10+1}/{-(-len(grupo)//_MAX_CITS_POR_CHAMADA)}")
                prompt = (
                    f"Você é um verificador de integridade acadêmica especializado em normas ABNT.\n\n"
                    f"{_REGRAS_ABNT}\n\n"
                    f"OBRA: {arquivo} (leitura: {modo})\n\n"
                    f"CONTEÚDO DA OBRA:\n\"\"\"{fonte_txt}\"\"\"\n\n"
                    f"Verifique as {len(sub)} citações desta obra encontradas na dissertação:\n\n"
                    f"{lista_str}\n\n{instrucao_final}"
                )
                content = [{"type": "text", "text": prompt}]

            elif arq_path and arq_path.suffix.lower() == ".pdf" and arq_path.exists():
                blocos_doc = _pdf_para_blocos(arq_path, sub_contextos)
                if not blocos_doc:
                    resultados += _erro_lista(sub, "PDF não pôde ser enviado à API")
                    continue
                modo = f"PDF nativo {len(blocos_doc)} bloco(s)"
                _log(f"     → PDF nativo {len(blocos_doc)} bloco(s) — lote {inicio//10+1}")
                prompt = (
                    f"Você é um verificador de integridade acadêmica especializado em normas ABNT.\n\n"
                    f"{_REGRAS_ABNT}\n\n"
                    f"OBRA: {arquivo} (leitura: {modo})\n"
                    f"O conteúdo está nos documentos anexados.\n\n"
                    f"Verifique as {len(sub)} citações desta obra encontradas na dissertação:\n\n"
                    f"{lista_str}\n\n{instrucao_final}"
                )
                content = blocos_doc + [{"type": "text", "text": prompt}]
            else:
                resultados += _erro_lista(sub, "Sem conteúdo disponível para verificação")
                continue

            resultados += _chamar_api(sub, content)

        except Exception as exc:
            resultados += _erro_lista(sub, str(exc))

        if inicio + _MAX_CITS_POR_CHAMADA < len(grupo):
            time.sleep(3)  # pausa entre lotes da mesma obra

    return resultados


# ── Quantos arquivos considerar como candidatos por citação ──────────────────
_MAX_CANDIDATOS_SUGESTAO = 5
_JANELA_CANDIDATO        = 8_000   # chars por candidato enviados ao Claude


def _melhor_pagina(palavras: list, paginas: list) -> Optional[int]:
    """Retorna o número (1-based) da página de `paginas` com mais palavras-chave
    em comum com `palavras`, ou None se nenhuma página tiver correspondência.
    Só faz sentido para PDF — único formato com paginação fixa gravada no arquivo."""
    melhor_idx, melhor_sc = None, 0
    for i, txt in enumerate(paginas):
        tl = txt.lower()
        sc = sum(1 for w in palavras if w in tl)
        if sc > melhor_sc:
            melhor_idx, melhor_sc = i, sc
    return (melhor_idx + 1) if melhor_idx is not None else None


def _candidatos_por_relevancia(contexto: str, material: dict, material_paginas: dict = None,
                                top_n: int = _MAX_CANDIDATOS_SUGESTAO) -> list:
    """Retorna [(nome_arquivo, pagina_ou_None, trecho_relevante)] para os top_n
    arquivos mais relevantes para o contexto da citação, com base em
    sobreposição de keywords. `pagina` só é preenchida para PDFs (únicos com
    paginação fixa gravada no arquivo — .docx não tem como apontar página real)."""
    material_paginas = material_paginas or {}
    palavras = [p for p in re.findall(r"[a-záéíóúâêîôûãõàç]{4,}", contexto.lower())
                if p not in _STOPWORDS]
    if not palavras:
        return []

    scored = []
    for nome, texto in material.items():
        if not texto:
            continue
        trecho = _trecho_relevante(contexto, texto, janela=_JANELA_CANDIDATO)
        sc = sum(1 for w in palavras if w in trecho.lower())
        if sc > 0:
            paginas = material_paginas.get(nome)
            pagina = _melhor_pagina(palavras, paginas) if paginas else None
            # Quando há página, restringe o trecho enviado ao texto dessa
            # página específica — garante que o que o Claude vê (e cita)
            # corresponde exatamente à página apontada.
            if pagina is not None:
                trecho = paginas[pagina - 1][:_JANELA_CANDIDATO]
            scored.append((sc, nome, pagina, trecho))

    scored.sort(key=lambda x: -x[0])
    return [(nome, pagina, trecho) for _, nome, pagina, trecho in scored[:top_n]]


def sugerir_obras_alternativas(cit: Citacao, material: dict, material_paginas: dict,
                                api_key: str, log_fn=None) -> dict:
    """
    Para uma citação sem fonte encontrada, busca entre todos os arquivos
    disponíveis aquele(s) que suportam semanticamente o mesmo argumento.
    Retorna dict com lista de sugestões ou indicação de nenhuma encontrada.
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)

    candidatos = _candidatos_por_relevancia(cit.contexto, material, material_paginas)
    if not candidatos:
        return {"encontrou": False, "sugestoes": [], "nota": "Nenhum candidato encontrado"}

    # Mapas auxiliares para validar a página devolvida pelo Claude — nunca
    # confiamos cegamente: para um trabalho científico, a página apontada
    # precisa ser real e conferível, não estimada.
    pagina_por_obra = {nome: pagina for nome, pagina, _ in candidatos}
    texto_pagina_por_obra = {
        nome: material_paginas[nome][pagina - 1]
        for nome, pagina, _ in candidatos
        if pagina is not None and nome in material_paginas
    }

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key, max_retries=5)

        blocos_obras = ""
        for i, (nome, pagina, trecho) in enumerate(candidatos, 1):
            cabecalho = f"=== OBRA {i}: {nome}" + (f" — PÁGINA {pagina}" if pagina else "") + " ==="
            blocos_obras += f"\n\n{cabecalho}\n{trecho}"

        prompt = (
            "Você é um especialista em integridade acadêmica. "
            "Uma dissertação contém a seguinte citação cujo autor/obra NÃO consta nas referências:\n\n"
            f"CITAÇÃO: {cit.texto}\n"
            f"CONTEXTO NA DISSERTAÇÃO: {_trecho_em_torno(cit.contexto, cit.texto, 400)}\n\n"
            "Abaixo estão trechos de obras disponíveis na pasta de referências. "
            "Verifique se alguma dessas obras contém um argumento ou dado que "
            "poderia validar o que foi afirmado na citação:\n"
            f"{blocos_obras}\n\n"
            "REGRA SOBRE PÁGINA — leia com atenção, isto é um trabalho científico "
            "e a página precisa ser EXATA e conferível, nunca estimada:\n"
            "- Alguns cabeçalhos de obra trazem '— PÁGINA N' (obtido diretamente da "
            "paginação real do arquivo PDF). Se o trecho que sustenta o argumento "
            "vier de um bloco com esse cabeçalho, copie esse N exatamente em "
            "\"pagina\".\n"
            "- Se o cabeçalho NÃO trouxer página (arquivo .docx/.doc, que não tem "
            "paginação fixa gravada no documento), retorne \"pagina\": null. "
            "NUNCA estime, calcule ou invente um número de página nesse caso.\n\n"
            "Responda SOMENTE com JSON (sem markdown):\n"
            '{"encontrou": true|false, '
            '"sugestoes": [{"obra": "nome_do_arquivo", '
            '"pagina": <número inteiro da página indicada no cabeçalho do bloco usado, ou null>, '
            '"trecho": "trecho literal ≤200 chars que suporta o argumento", '
            '"citacao_sugerida": "como citar em ABNT ex: SOBRENOME (ANO) ou (SOBRENOME, ANO)", '
            '"referencia_completa": "referência bibliográfica completa em ABNT da obra, '
            'montada a partir das informações disponíveis no trecho/obra (autor, título, '
            'local, editora, ano); se faltar algum dado, omita-o — não invente", '
            '"justificativa": "1 frase explicando por que essa obra valida o argumento"}]}'
        )

        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1536,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = re.sub(r"^```(?:json)?", "", resp.content[0].text.strip()).strip("`").strip()
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            parsed = json.loads(m.group()) if m else {}

        parsed.setdefault("encontrou", False)
        parsed.setdefault("sugestoes", [])

        # Validação defensiva da página — nunca exibimos um número não conferível:
        # 1) se a obra não tem paginação real disponível (não-PDF), zera;
        # 2) se tem, exige que o trecho citado realmente apareça naquela página
        #    específica do arquivo — caso contrário, a página é descartada.
        _norm = lambda s: re.sub(r"\s+", " ", (s or "")).strip().lower()
        for sg in parsed["sugestoes"]:
            sg.setdefault("pagina", None)
            sg.setdefault("referencia_completa", "")
            nome = sg.get("obra", "")
            pagina = sg.get("pagina")
            if pagina is None:
                continue
            if pagina_por_obra.get(nome) != pagina:
                sg["pagina"] = None
                continue
            txt_pag = _norm(texto_pagina_por_obra.get(nome, ""))
            trecho_norm = _norm(sg.get("trecho", ""))[:80]
            if not trecho_norm or trecho_norm not in txt_pag:
                _log(f"         ⚠ Página {pagina} de {nome} não confirma o trecho citado — descartando número de página")
                sg["pagina"] = None

        return parsed

    except Exception as exc:
        _log(f"         ⚠ Erro na sugestão: {str(exc)[:120]}")
        return {"encontrou": False, "sugestoes": [], "nota": str(exc)[:200]}


def verificar_claude(cit: Citacao, ref: Referencia, arquivo: str,
                     arq_path: Optional[Path], texto_fonte: str,
                     api_key: str) -> dict:
    try:
        import anthropic
        # max_retries=5 ativa retry automático do SDK para 429 e 529
        client = anthropic.Anthropic(api_key=api_key, max_retries=5)

        ext = arq_path.suffix.lower() if arq_path else ""

        # Detecção de compêndio (padrão ABNT "In:") para instruir Claude
        eh_comp = bool(re.search(r"\bIn:\s*", ref.texto, re.IGNORECASE))
        info_comp = (
            f"ATENÇÃO — COMPÊNDIO: Este é um capítulo de livro. "
            f"Localize o capítulo de {ref.sobrenome} dentro do livro "
            f"e baseie a análise nesse capítulo específico."
            if eh_comp else ""
        )

        # ── PRIORIDADE 1: texto extraído (token-eficiente, evita rate limit) ──
        # Com PyMuPDF instalado, sempre preferimos texto. PDF nativo só como fallback.
        if texto_fonte:
            fonte_txt = _sanitizar(_trecho_relevante(cit.contexto, texto_fonte,
                                          janela=_LIMITE_TEXTO_SEMANTICA))
            modo = f"texto ({len(texto_fonte)//1024} KB total → seleção {len(fonte_txt)//1024} KB)"
            prompt_txt = _PROMPT_CLAUDE_TEXTO.format(
                regras_abnt=_REGRAS_ABNT,
                citacao=cit.texto,
                contexto=_trecho_em_torno(cit.contexto, cit.texto, 1200),
                referencia=ref.texto[:400],
                arquivo=arquivo,
                modo_leitura=modo,
                info_compendio=info_comp,
                fonte=fonte_txt,
            )
            content = [{"type": "text", "text": prompt_txt}]

        # ── PRIORIDADE 2: PDF nativo (fallback quando sem texto extraído) ──
        elif ext == ".pdf" and arq_path and arq_path.exists():
            blocos_doc = _pdf_para_blocos(arq_path, cit.contexto)
            if not blocos_doc:
                return {"veredicto": "SEM_FONTE",
                        "justificativa": "PDF não pôde ser enviado à API (muito grande ou corrompido).",
                        "trecho_fonte": ""}
            modo = f"PDF nativo — {len(blocos_doc)} bloco(s)"
            prompt_txt = _PROMPT_CLAUDE.format(
                regras_abnt=_REGRAS_ABNT,
                citacao=cit.texto,
                contexto=_trecho_em_torno(cit.contexto, cit.texto, 1200),
                referencia=ref.texto[:400],
                arquivo=arquivo,
                modo_leitura=modo,
                info_compendio=info_comp,
            )
            content = blocos_doc + [{"type": "text", "text": prompt_txt}]

        else:
            return {"veredicto": "SEM_FONTE",
                    "justificativa": "Nenhum conteúdo disponível para verificação.",
                    "trecho_fonte": ""}

        # SDK já faz retry automático (max_retries=5) para 429/529
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": content}],
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
    sugestoes   = resultado.get("sugestoes_alternativas", [])

    corretas   = [v for v in verifics if v.get("veredicto") == "CORRETO"]
    incorretas = [v for v in verifics if v.get("veredicto") == "INCORRETO"]
    parciais   = [v for v in verifics if v.get("veredicto") == "PARCIAL"]
    # "SEM_FONTE" cobre DUAS situações distintas que não podem ser misturadas
    # sob o mesmo rótulo — uma diz que não há arquivo, a outra que HÁ arquivo
    # mas seu conteúdo não confirma a citação (a IA já indica a referência e
    # o arquivo encontrado, então rotular como "arquivo não encontrado" é
    # contraditório e confunde o usuário):
    #   (a) arquivo_fonte vazio  → a obra citada não está na pasta de apoio
    #   (b) arquivo_fonte preenchido → obra localizada, mas seu conteúdo (na
    #       avaliação da IA) não sustenta o argumento da citação
    sem_fonte_total = [v for v in verifics if v.get("veredicto") == "SEM_FONTE"]
    sem_arquivo     = [v for v in sem_fonte_total if not v.get("arquivo_fonte")]
    sem_conteudo    = [v for v in sem_fonte_total if v.get("arquivo_fonte")]
    sugs_com_match = [s for s in sugestoes if s.get("encontrou") and s.get("sugestoes")]

    # Desvios de FORMATAÇÃO ABNT — independentes do veredicto semântico acima;
    # uma citação pode estar "CORRETA" no argumento e ainda assim fora do padrão ABNT.
    desvios_formato = [v for v in verifics
                       if not v.get("formato_abnt", {}).get("conforme", True)]

    def e(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def secao(titulo, items, cor, fn):
        if not items:
            return ""
        rows = "".join(f"<li>{fn(i)}</li>" for i in items)
        return f'<h2 style="color:{cor}">{e(titulo)} ({len(items)})</h2><ul>{rows}</ul>'

    def card_formato(v):
        obs = v.get("formato_abnt", {}).get("observacao", "")
        return (f'<b>{e(v["citacao"])}</b>'
                f'<br><small>Desvio identificado: {e(obs) if obs else "—"}</small>'
                f'<details><summary><small>ver contexto</small></summary>'
                f'<small>{e(v.get("contexto","")[:300])}</small></details>')

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
.sug-obra {{background:#fff9e6;border:1px solid #f0c040;border-radius:6px;padding:.6em 1em;margin:.4em 0}}
.sug-obra b {{color:#7d5a00}}
.sug-citacao {{font-family:monospace;background:#f5f5f5;padding:.2em .5em;border-radius:3px;color:#1a2744}}
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
  <div class="card" style="border-color:#8e7c20"><b>{len(desvios_formato)}</b>Desvios de<br>formatação ABNT</div>
  <div class="card" style="border-color:#f0c040"><b>{len(sugs_com_match)}</b>Obras<br>sugeridas</div>
</div>
"""

    html += secao("⚠ Citações SEM referência na lista bibliográfica", citadas_sr, "#e74c3c",
        lambda c: (f'<b class="err">{e(c.texto)}</b>'
                   f'<details><summary><small>ver contexto</small></summary>'
                   f'<small>{e(_trecho_em_torno(c.contexto, c.texto, 350))}</small></details>'))

    html += secao("⚠ Referências listadas SEM citação no texto", refs_sc, "#e67e22",
        lambda r: f'<span class="warn">{e(r.texto[:250])}</span>')

    if incorretas:
        html += secao("✗ INCORRETAS — argumento diverge da obra (problema SEMÂNTICO)", incorretas, "#c0392b", card)
    if parciais:
        html += secao("~ PARCIAIS — citação com ressalvas no argumento (problema SEMÂNTICO)", parciais, "#d35400", card)
    if corretas:
        html += secao("✓ CORRETAS — argumento confirmado pela obra", corretas, "#27ae60", card)
    if sem_arquivo:
        html += secao("? ARQUIVO NÃO LOCALIZADO — a obra citada não está na pasta de apoio",
                      sem_arquivo, "#7f8c8d", card)
    if sem_conteudo:
        html += secao("⚠ ARGUMENTO NÃO CONFIRMADO NA OBRA — arquivo localizado, "
                      "mas seu conteúdo não sustenta a citação (problema SEMÂNTICO)",
                      sem_conteudo, "#a04000", card)
        html += ('<p><em>A obra referenciada foi encontrada na pasta de apoio, mas a IA '
                 'não localizou nela o argumento ou dado citado na dissertação. Veja '
                 'mais abaixo, em "OBRAS SUGERIDAS", se uma fonte alternativa que '
                 'sustenta o mesmo argumento foi localizada.</em></p>')

    if desvios_formato:
        html += secao(
            "📐 DESVIOS DE FORMATAÇÃO ABNT — problema de FORMA, não de conteúdo",
            desvios_formato, "#8e7c20", card_formato)
        html += ('<p><em>As citações abaixo não seguem rigorosamente o padrão ABNT '
                 '(ex.: pontuação, parênteses, espaçamento), mas isso NÃO significa '
                 'que o argumento esteja incorreto — a fidelidade ao conteúdo da obra '
                 'é avaliada separadamente, nas seções de divergência semântica acima.</em></p>')

    _MOTIVO_SUGESTAO = {
        "INCORRETO": "O argumento diverge do que a obra citada realmente afirma.",
        "PARCIAL":   "O argumento só é parcialmente sustentado pela obra citada.",
        "SEM_FONTE": ("A obra citada não foi localizada na pasta de apoio."),
    }

    def _motivo_sugestao(s):
        origem = s.get("veredicto_origem", "")
        if origem == "SEM_FONTE" and s.get("arquivo_origem"):
            return ("O arquivo da obra citada foi localizado, mas seu conteúdo "
                    "não confirma o argumento da citação.")
        return _MOTIVO_SUGESTAO.get(origem, "")

    if sugestoes:
        com_match = [s for s in sugestoes if s.get("encontrou") and s.get("sugestoes")]
        sem_match = [s for s in sugestoes if not (s.get("encontrou") and s.get("sugestoes"))]
        if com_match:
            html += f'<h2 style="color:#b8860b">💡 FONTES ALTERNATIVAS ENCONTRADAS — citações com argumento não confirmado pela obra citada, mas sustentado por outra obra disponível ({len(com_match)})</h2>'
            html += ('<p><em>Nos casos abaixo, o argumento da dissertação não foi confirmado na obra '
                     'indicada como referência (seja por divergência de conteúdo, seja por a obra não '
                     'ter sido localizada), mas outra obra do material de apoio sustenta o mesmo argumento. '
                     'Considere substituir a citação/referência pelo indicado.</em></p><ul>')
            for s in com_match:
                motivo = _motivo_sugestao(s)
                ref_origem = s.get("referencia_origem", "")
                html += f'<li style="border-left:5px solid #f0c040;background:#fffdf0;padding:.7em 1em;margin:.5em 0;border-radius:0 6px 6px 0">'
                html += f'<b>Citação original:</b> {e(s["citacao"])}<br>'
                if ref_origem:
                    html += f'<small><b>Referência atualmente indicada:</b> {e(ref_origem[:150])}</small><br>'
                if motivo:
                    html += f'<small style="color:#a04000"><b>Por que buscamos uma alternativa:</b> {e(motivo)}</small><br>'
                html += f'<small style="color:#555">{e(s["contexto"][:300])}</small><br><br>'
                for sg in s["sugestoes"]:
                    pagina = sg.get("pagina")
                    if pagina:
                        pag_html = f'<b>Página:</b> {e(str(pagina))} <small style="color:#888">(localização exata, extraída do PDF — conferível)</small><br>'
                    else:
                        pag_html = ('<b>Página:</b> <small style="color:#888">não disponível — '
                                    'arquivo sem paginação fixa gravada (.docx) ou trecho não confirmado '
                                    'na página apontada; não estimamos números de página</small><br>')
                    ref_completa = sg.get("referencia_completa", "")
                    ref_html = (f'<b>Referência completa (ABNT):</b> {e(ref_completa)}<br>'
                                if ref_completa else "")
                    html += (
                        f'<div class="sug-obra">'
                        f'<b>Arquivo:</b> {e(sg.get("obra", "?"))}<br>'
                        f'{pag_html}'
                        f'<b>Citação sugerida:</b> <span class="sug-citacao">{e(sg.get("citacao_sugerida", "?"))}</span><br>'
                        f'{ref_html}'
                        f'<b>Justificativa:</b> {e(sg.get("justificativa", ""))}<br>'
                    )
                    if sg.get("trecho"):
                        html += f'<blockquote>"{e(sg["trecho"])}"</blockquote>'
                    html += '</div>'
                html += '</li>'
            html += '</ul>'
        if sem_match:
            html += f'<h2 style="color:#95a5a6">— SEM SUGESTÃO — citações cujo argumento não foi encontrado nas obras disponíveis ({len(sem_match)})</h2><ul>'
            for s in sem_match:
                html += (f'<li class="info"><b>{e(s["citacao"])}</b>'
                         f'<details><summary><small>ver contexto</small></summary>'
                         f'<small>{e(s["contexto"][:300])}</small></details></li>')
            html += '</ul>'

    if not verifics and pareamentos:
        html += f'<h2 style="color:#3498db">✓ Citações pareadas com referências ({len(pareamentos)})</h2>'
        html += '<p><em>Sem verificação semântica — informe a chave de API Claude para verificar a coerência dos argumentos.</em></p><ul>'
        for c, r in pareamentos:
            html += (f'<li class="info"><b>{e(c.texto)}</b>'
                     f'<br><small>Ref.: {e(r.texto[:180])}</small>'
                     f'<details><summary><small>ver contexto</small></summary>'
                     f'<small>{e(_trecho_em_torno(c.contexto, c.texto, 300))}</small></details></li>')
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
    L("VERSÃO: 2026-06-07-v8.17")
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
    try:
        diss_resolvido = Path(diss_path).resolve()
    except OSError:
        diss_resolvido = None
    arquivos = sorted(
        [f for f in Path(refs_dir).rglob("*")
         if f.is_file() and f.suffix.lower() in extensoes
         and (diss_resolvido is None or f.resolve() != diss_resolvido)],
        key=lambda f: f.name.lower()
    )
    L(f"  {len(arquivos)} arquivo(s) encontrado(s) na pasta")
    L("")

    material         = {}   # nome → texto extraído  (para matching por conteúdo)
    material_paths   = {}   # nome → Path            (para envio nativo ao Claude)
    material_paginas = {}   # nome → [texto_pág_1, texto_pág_2, ...]  (só PDF — única
                            # fonte com paginação fixa gravada no arquivo; permite
                            # apontar o número exato da página de um trecho sugerido)
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
            if is_pdf:
                paginas = ler_pdf_paginas(arq)
                if paginas:
                    material_paginas[arq.name] = paginas
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

            # ── Registra citações SEM arquivo encontrado ─────────────────
            sem_arq = [(cit, ref, arq, arq_path, txt)
                       for cit, ref, arq, arq_path, txt in candidatos_semantica
                       if not arq]
            for cit, ref, arq, arq_path, txt in sem_arq:
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
                    "formato_abnt": {"conforme": True, "observacao": ""},
                })

            # ── Agrupa citações por arquivo fonte ─────────────────────────
            from collections import defaultdict as _ddict
            grupos: dict = _ddict(list)
            for item in candidatos_semantica:
                cit, ref, arq, arq_path, txt = item
                if arq:
                    grupos[arq].append(item)

            total_grupos = len(grupos)
            L(f"  → {total_grupos} obra(s) distinta(s) a verificar")
            L("")

            for gi, (arq_nome, grupo) in enumerate(sorted(grupos.items()), 1):
                if stop_fn and stop_fn():
                    L("\n[INTERROMPIDO]")
                    break

                _, _, _, arq_path_g, txt_g = grupo[0]
                tam = f"{arq_path_g.stat().st_size // 1024} KB" if arq_path_g else "—"
                modo_txt = f"texto {len(txt_g):,} chars" if txt_g else "somente PDF nativo"
                L(f"  [{gi:>3}/{total_grupos}] {arq_nome}  ({tam}, {modo_txt})")
                L(f"           {len(grupo)} citação(ões) desta obra")

                resultados_g = verificar_grupo_claude(
                    grupo, arq_nome, arq_path_g, txt_g, api_key, log_fn=L
                )

                for (cit, ref, arq, arq_path, txt), v in zip(grupo, resultados_g):
                    emoji = {"CORRETO": "✓", "INCORRETO": "✗", "PARCIAL": "⚠",
                             "SEM_FONTE": "?", "ERRO": "!"}.get(v.get("veredicto", ""), "?")
                    L(f"           {emoji} {cit.texto[:60]}  →  {v.get('veredicto')} — {v.get('justificativa','')[:70]}")

                    ref_obj = ref if ref is not None else Referencia(
                        texto=f"{cit.autores[0]} ({cit.ano}) — sem entrada na lista de referências",
                        sobrenome=cit.autores[0],
                        sobrenomes=cit.autores,
                        ano=cit.ano,
                        titulo=" ".join(cit.autores),
                    )
                    verificacoes.append({
                        "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                        "contexto": cit.contexto,
                        "referencia": ref_obj.texto,
                        "arquivo_fonte": arq, **v,
                    })

                # Pausa entre obras diferentes
                if gi < total_grupos:
                    time.sleep(3)

    # ────────────────────────────────────────────────────────────────────
    # ETAPA 5 — Sugestão de obras alternativas para citações sem fonte
    # ────────────────────────────────────────────────────────────────────
    sugestoes_alternativas = []
    L("")
    L("=" * 65)
    L("ETAPA 5 — Buscando obras alternativas para citações com argumento não confirmado")
    L("=" * 65)

    # Busca uma fonte alternativa sempre que o argumento da citação NÃO esteja
    # confirmado pela obra que ela referencia — seja porque (a) a referência
    # não existe/o arquivo não foi localizado (SEM_FONTE sem arquivo_fonte),
    # (b) o arquivo foi localizado mas seu conteúdo não confirma o argumento
    # (SEM_FONTE com arquivo_fonte), ou (c) o argumento diverge do que a obra
    # realmente afirma (INCORRETO/PARCIAL). Em todos esses casos, vale a pena
    # apontar — se existir — uma obra do material de apoio que sustente o
    # mesmo argumento, com página, citação e referência corretas para troca
    # na dissertação.
    candidatos_sugestao = [v for v in verificacoes
                           if v.get("veredicto") in ("SEM_FONTE", "INCORRETO", "PARCIAL")]
    material_com_texto = {n: t for n, t in material.items() if t}

    if sem_verificacao or not api_key.strip():
        L("  [Etapa ignorada — verificação semântica desativada ou sem chave API]")
    elif not material_com_texto:
        L("  [Sem arquivos com texto extraído — não é possível busca semântica]")
    elif not candidatos_sugestao:
        L("  ✓ Nenhuma citação com argumento não confirmado requer sugestão")
    else:
        L(f"  → {len(candidatos_sugestao)} citação(ões) com argumento não confirmado a analisar")
        L(f"  → {len(material_com_texto)} obra(s) disponíveis para busca semântica")
        L("")
        for i, v in enumerate(candidatos_sugestao, 1):
            if stop_fn and stop_fn():
                L("\n[INTERROMPIDO]")
                break
            cit_txt = v["citacao"]
            cit_ctx = v["contexto"]
            L(f"  [{i:>3}/{len(candidatos_sugestao)}] {cit_txt[:70]}  (veredicto: {v.get('veredicto')})")
            cit_obj = Citacao(
                texto=cit_txt,
                autores=v.get("autores", []),
                ano=v.get("ano", ""),
                pagina=None,
                contexto=cit_ctx,
                paragrafo_idx=0,
            )
            resultado_sug = sugerir_obras_alternativas(cit_obj, material_com_texto, material_paginas, api_key, L)
            if resultado_sug.get("encontrou") and resultado_sug.get("sugestoes"):
                for s in resultado_sug["sugestoes"]:
                    pag_txt = f", p. {s['pagina']}" if s.get("pagina") else ""
                    L(f"         ✔ Obra sugerida: {s.get('obra','?')}{pag_txt} → {s.get('citacao_sugerida','?')}")
            else:
                L(f"         — Nenhuma obra de suporte encontrada")
            sugestoes_alternativas.append({
                "citacao":          cit_txt,
                "contexto":         cit_ctx,
                "autores":          v.get("autores", []),
                "ano":              v.get("ano", ""),
                "veredicto_origem": v.get("veredicto", ""),
                "referencia_origem": v.get("referencia", ""),
                "arquivo_origem":   v.get("arquivo_fonte", ""),
                **resultado_sug,
            })
            if i < len(candidatos_sugestao):
                time.sleep(3)

    return {
        "citacoes":               citacoes,
        "referencias":            referencias,
        "pareamentos":            pareamentos,
        "citadas_sem_ref":        citadas_sem_ref,
        "refs_sem_citacao":       refs_sem_citacao,
        "verificacoes":           verificacoes,
        "sugestoes_alternativas": sugestoes_alternativas,
        "diss_path":              diss_path,
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

    log("Bem-vindo ao Verificador de Citações Acadêmicas  [versão 2025-06-02-v8.4]")
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
