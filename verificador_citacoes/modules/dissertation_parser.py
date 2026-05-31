"""
Extrai citações in-text e entradas da lista de referências de uma dissertação .docx.
Suporta o padrão ABNT NBR 10520 (citações) e ABNT NBR 6023 (referências).
"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import docx


# ──────────────────────────────────────────────
# Estruturas de dados
# ──────────────────────────────────────────────

@dataclass
class CitacaoInText:
    """Uma citação encontrada no corpo da dissertação."""
    texto_original: str          # ex.: "(SILVA, 2020, p. 45)"
    autores: list[str]           # ex.: ["SILVA"]
    ano: str                     # ex.: "2020"
    pagina: Optional[str]        # ex.: "45" ou "45-46"
    contexto: str                # ±300 chars ao redor da citação
    paragrafo_idx: int           # posição do parágrafo na dissertação
    chave: str = ""              # "SILVA_2020" — preenchido após parsing

    def __post_init__(self):
        autores_str = "; ".join(self.autores)
        self.chave = f"{autores_str}_{self.ano}"


@dataclass
class EntradaReferencia:
    """Uma entrada da lista de referências ABNT."""
    texto_completo: str          # linha original da referência
    autor_sobrenome: str         # primeiro sobrenome extraído
    ano: str                     # ano da publicação
    titulo: str                  # título identificado
    chave: str = ""              # "SILVA_2020"

    def __post_init__(self):
        self.chave = f"{self.autor_sobrenome}_{self.ano}"


# ──────────────────────────────────────────────
# Padrões ABNT
# ──────────────────────────────────────────────

# Padrões para citações in-text
_PAD_SOBRENOME = r"[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç'\-]+"
_PAD_ANO       = r"\d{4}[a-z]?"
_PAD_PAGINA    = r"(?:,\s*p(?:p)?\.?\s*([\d\-–]+))?"

# "(SILVA, 2020)" ou "(SILVA, 2020, p. 45)" ou "(SILVA et al., 2020)"
_RE_PAREN = re.compile(
    r'\(('
    r'(?:' + _PAD_SOBRENOME + r'(?:\s+' + _PAD_SOBRENOME + r')*'
    r'(?:\s+et\s+al\.?)?'
    r'(?:;\s*' + _PAD_SOBRENOME + r'(?:\s+' + _PAD_SOBRENOME + r')*'
    r'(?:\s+et\s+al\.?)?)*)'
    r'),\s*(' + _PAD_ANO + r')' + _PAD_PAGINA + r'\)',
    re.IGNORECASE
)

# "SILVA (2020)" ou "Silva (2020, p. 45)"
_RE_NARRATIVO = re.compile(
    r'(' + _PAD_SOBRENOME + r'(?:;\s*' + _PAD_SOBRENOME + r')*'
    r'(?:,?\s*et\s+al\.?)?)\s+\((' + _PAD_ANO + r')' + _PAD_PAGINA + r'\)',
)

# Marcadores de início da seção de referências
_RE_INICIO_REFS = re.compile(
    r'^\s*REFER[EÊ]NCIAS\s*(BIBLIOGR[ÁA]FICAS?)?\s*$',
    re.IGNORECASE
)

# Padrão de uma entrada de referência ABNT:
# SOBRENOME, Nome[s]. Título... ano.
_RE_ENTRADA_REF = re.compile(
    r'^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+)'  # sobrenome
    r'(?:,\s*.+?)?[.;]\s*'       # demais autores (opcional)
    r'(.+?)[.;]\s*'               # título (captura aproximada)
    r'.*?(\d{4})',                # ano
    re.DOTALL
)


# ──────────────────────────────────────────────
# Funções auxiliares
# ──────────────────────────────────────────────

def _extrair_contexto(paragrafos: list[str], idx: int, janela: int = 300) -> str:
    """Retorna até `janela` chars antes e depois do parágrafo."""
    texto_antes = " ".join(paragrafos[max(0, idx-2):idx])
    texto_depois = " ".join(paragrafos[idx+1:idx+3])
    para = paragrafos[idx]
    return (texto_antes[-janela:] + " " + para + " " + texto_depois[:janela]).strip()


def _normalizar_autores(raw: str) -> list[str]:
    """Divide uma string de autores em lista de sobrenomes normalizados."""
    # Remove "et al." e divide por ";"
    raw = re.sub(r',?\s*et\s+al\.?', '', raw, flags=re.IGNORECASE)
    partes = [p.strip() for p in re.split(r';', raw) if p.strip()]
    # Pega apenas o primeiro token (sobrenome) de cada parte
    sobrenomes = []
    for p in partes:
        tokens = p.split()
        if tokens:
            sobrenomes.append(tokens[0].upper())
    return sobrenomes


# ──────────────────────────────────────────────
# Parser principal
# ──────────────────────────────────────────────

class DissertationParser:
    def __init__(self, caminho: str | Path):
        self.caminho = Path(caminho)
        self._paragrafos: list[str] = []
        self._idx_inicio_refs: int = -1
        self.citacoes: list[CitacaoInText] = []
        self.referencias: list[EntradaReferencia] = []

    # ── Leitura ──────────────────────────────

    def carregar(self) -> "DissertationParser":
        doc = docx.Document(str(self.caminho))
        self._paragrafos = [p.text for p in doc.paragraphs]
        self._encontrar_inicio_referencias()
        return self

    def _encontrar_inicio_referencias(self):
        for i, texto in enumerate(self._paragrafos):
            if _RE_INICIO_REFS.match(texto.strip()):
                self._idx_inicio_refs = i
                break

    # ── Citações in-text ─────────────────────

    def extrair_citacoes(self) -> list[CitacaoInText]:
        """Varre o corpo (antes das referências) e coleta todas as citações."""
        limite = self._idx_inicio_refs if self._idx_inicio_refs > 0 else len(self._paragrafos)
        textos = self._paragrafos[:limite]
        citacoes: list[CitacaoInText] = []
        vistos: set[tuple] = set()

        for idx, para in enumerate(textos):
            contexto = _extrair_contexto(textos, idx)

            # Padrão parentético: (SILVA, 2020, p. 45)
            for m in _RE_PAREN.finditer(para):
                autores = _normalizar_autores(m.group(1))
                ano = m.group(2)
                pagina = m.group(3) if m.lastindex >= 3 else None
                chave_vis = (tuple(autores), ano, pagina, idx)
                if chave_vis not in vistos:
                    vistos.add(chave_vis)
                    citacoes.append(CitacaoInText(
                        texto_original=m.group(0),
                        autores=autores,
                        ano=ano,
                        pagina=pagina,
                        contexto=contexto,
                        paragrafo_idx=idx,
                    ))

            # Padrão narrativo: Silva (2020)
            for m in _RE_NARRATIVO.finditer(para):
                autores = _normalizar_autores(m.group(1))
                ano = m.group(2)
                pagina = m.group(3) if m.lastindex >= 3 else None
                chave_vis = (tuple(autores), ano, pagina, idx)
                if chave_vis not in vistos:
                    vistos.add(chave_vis)
                    citacoes.append(CitacaoInText(
                        texto_original=m.group(0),
                        autores=autores,
                        ano=ano,
                        pagina=pagina,
                        contexto=contexto,
                        paragrafo_idx=idx,
                    ))

        self.citacoes = citacoes
        return citacoes

    # ── Lista de referências ─────────────────

    def extrair_referencias(self) -> list[EntradaReferencia]:
        """Extrai entradas da seção REFERÊNCIAS."""
        if self._idx_inicio_refs < 0:
            print("[AVISO] Seção REFERÊNCIAS não encontrada. Verifique o título da seção.")
            return []

        entradas: list[EntradaReferencia] = []
        buffer_linhas: list[str] = []

        def _processar_buffer(buf: list[str]) -> Optional[EntradaReferencia]:
            texto = " ".join(buf).strip()
            if not texto:
                return None
            m = _RE_ENTRADA_REF.match(texto)
            if m:
                sobrenome = m.group(1).strip().upper()
                titulo = m.group(2).strip()
                ano = m.group(3).strip()
                return EntradaReferencia(
                    texto_completo=texto,
                    autor_sobrenome=sobrenome,
                    ano=ano,
                    titulo=titulo,
                )
            # Tentativa menos rígida: extrai ano de qualquer jeito
            m_ano = re.search(r'\b(\d{4})\b', texto)
            m_sob = re.match(r'^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+)', texto)
            if m_sob and m_ano:
                return EntradaReferencia(
                    texto_completo=texto,
                    autor_sobrenome=m_sob.group(1).upper(),
                    ano=m_ano.group(1),
                    titulo=texto[:80],
                )
            return None

        paras_refs = self._paragrafos[self._idx_inicio_refs + 1:]
        _RE_NOVA_ENTRADA = re.compile(
            r'^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+[,;]'
        )

        for linha in paras_refs:
            linha = linha.strip()
            if not linha:
                if buffer_linhas:
                    entrada = _processar_buffer(buffer_linhas)
                    if entrada:
                        entradas.append(entrada)
                    buffer_linhas = []
                continue

            if _RE_NOVA_ENTRADA.match(linha) and buffer_linhas:
                entrada = _processar_buffer(buffer_linhas)
                if entrada:
                    entradas.append(entrada)
                buffer_linhas = [linha]
            else:
                buffer_linhas.append(linha)

        if buffer_linhas:
            entrada = _processar_buffer(buffer_linhas)
            if entrada:
                entradas.append(entrada)

        self.referencias = entradas
        return entradas

    # ── Relatório de cruzamento ───────────────

    def cruzar_citacoes_referencias(self) -> dict:
        """
        Retorna:
          citadas_sem_referencia: citações sem entrada correspondente na lista
          referenciadas_sem_citacao: entradas sem citação correspondente no texto
          pareamentos: lista de (citacao, referencia) que fazem par
        """
        chaves_citadas = {(sob, c.ano) for c in self.citacoes for sob in c.autores}
        chaves_refs    = {(r.autor_sobrenome, r.ano): r for r in self.referencias}

        citadas_sem_ref: list[CitacaoInText] = []
        pareamentos: list[tuple[CitacaoInText, EntradaReferencia]] = []

        for cit in self.citacoes:
            encontrou = False
            for sob in cit.autores:
                chave = (sob, cit.ano)
                if chave in chaves_refs:
                    pareamentos.append((cit, chaves_refs[chave]))
                    encontrou = True
                    break
            if not encontrou:
                citadas_sem_ref.append(cit)

        chaves_citadas_set = {chave for chave in chaves_citadas}
        refs_sem_citacao = [
            r for (sob, ano), r in chaves_refs.items()
            if (sob, ano) not in chaves_citadas_set
        ]

        return {
            "citadas_sem_referencia": citadas_sem_ref,
            "referenciadas_sem_citacao": refs_sem_citacao,
            "pareamentos": pareamentos,
        }
