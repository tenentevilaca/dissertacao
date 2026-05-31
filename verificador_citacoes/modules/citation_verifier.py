"""
Usa a API Claude para verificar se uma citação realmente reflete
o conteúdo da obra referenciada.
"""

from dataclasses import dataclass
from enum import Enum
from typing import Optional

import anthropic

from config import (
    ANTHROPIC_API_KEY,
    CLAUDE_MODEL,
    MAX_TOKENS_VERIFICACAO,
    MAX_CHARS_CONTEXTO_DISSERTACAO,
    MAX_CHARS_CONTEXTO_FONTE,
)
from modules.dissertation_parser import CitacaoInText, EntradaReferencia
from modules.reference_reader import DocumentoReferencia, ReferenceReader


class Veredicto(str, Enum):
    CORRETO      = "CORRETO"       # citação fiel ao conteúdo da fonte
    INCORRETO    = "INCORRETO"     # citação distorce ou contradiz a fonte
    PARCIAL      = "PARCIAL"       # parcialmente suportada, mas com ressalvas
    SEM_FONTE    = "SEM_FONTE"     # documento fonte não foi encontrado na pasta
    ERRO         = "ERRO"          # falha na chamada de API


@dataclass
class ResultadoVerificacao:
    citacao: CitacaoInText
    referencia: Optional[EntradaReferencia]
    documento_fonte: Optional[DocumentoReferencia]
    veredicto: Veredicto
    justificativa: str
    trecho_fonte_usado: str = ""


_SYSTEM_PROMPT = """Você é um assistente especializado em verificação de integridade
acadêmica. Sua tarefa é avaliar se uma citação feita em uma dissertação acadêmica
representa fielmente o conteúdo da obra referenciada.

Responda SEMPRE em português brasileiro e siga o formato exato abaixo:

VEREDICTO: <CORRETO | INCORRETO | PARCIAL>
JUSTIFICATIVA: <uma única frase clara explicando o motivo>
"""

_USER_TEMPLATE = """## Contexto da dissertação (trecho ao redor da citação)

{contexto_dissertacao}

---
## Citação em análise

{citacao_original}

---
## Referência bibliográfica

{referencia_completa}

---
## Trecho do documento-fonte encontrado

{trecho_fonte}

---
## Pergunta

A citação {citacao_original} representa fielmente o que o documento-fonte diz?
Considere o contexto em que a citação aparece na dissertação e o trecho do documento-fonte acima.
"""


class CitationVerifier:
    def __init__(self, reader: ReferenceReader):
        self.reader = reader
        self._client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        # Cache de system prompt (reutilizado em todas as chamadas)
        self._cached_system: list[dict] = [
            {
                "type": "text",
                "text": _SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ]

    def verificar(
        self,
        citacao: CitacaoInText,
        referencia: EntradaReferencia,
    ) -> ResultadoVerificacao:
        """Verifica uma única citação contra a referência e o documento fonte."""

        # Palavras do título para busca
        palavras_titulo = referencia.titulo.split() if referencia.titulo else []

        # Busca o documento fonte
        candidatos = self.reader.buscar_documento(
            sobrenome=referencia.autor_sobrenome,
            ano=referencia.ano,
            titulo_palavras=palavras_titulo,
        )

        if not candidatos:
            return ResultadoVerificacao(
                citacao=citacao,
                referencia=referencia,
                documento_fonte=None,
                veredicto=Veredicto.SEM_FONTE,
                justificativa="Nenhum arquivo correspondente encontrado na pasta de referências.",
            )

        doc_fonte = candidatos[0]
        trecho = self.reader.trecho_por_pagina(
            doc_fonte,
            citacao.pagina,
            janela=MAX_CHARS_CONTEXTO_FONTE,
        )

        prompt_user = _USER_TEMPLATE.format(
            contexto_dissertacao=citacao.contexto[:MAX_CHARS_CONTEXTO_DISSERTACAO],
            citacao_original=citacao.texto_original,
            referencia_completa=referencia.texto_completo,
            trecho_fonte=trecho,
        )

        try:
            resposta = self._client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=MAX_TOKENS_VERIFICACAO,
                system=self._cached_system,
                messages=[{"role": "user", "content": prompt_user}],
            )
            texto_resposta = resposta.content[0].text.strip()
            veredicto, justificativa = _parsear_resposta(texto_resposta)
        except anthropic.APIError as e:
            veredicto = Veredicto.ERRO
            justificativa = f"Erro na API: {e}"
            trecho = ""

        return ResultadoVerificacao(
            citacao=citacao,
            referencia=referencia,
            documento_fonte=doc_fonte,
            veredicto=veredicto,
            justificativa=justificativa,
            trecho_fonte_usado=trecho[:500] + "..." if len(trecho) > 500 else trecho,
        )

    def verificar_lote(
        self,
        pareamentos: list[tuple[CitacaoInText, EntradaReferencia]],
        verbose: bool = True,
    ) -> list[ResultadoVerificacao]:
        """Verifica todos os pareamentos (citação, referência) em sequência."""
        resultados = []
        total = len(pareamentos)
        for i, (cit, ref) in enumerate(pareamentos, start=1):
            if verbose:
                print(f"  [{i}/{total}] Verificando: {cit.texto_original[:60]}…")
            resultado = self.verificar(cit, ref)
            resultados.append(resultado)
        return resultados


def _parsear_resposta(texto: str) -> tuple[Veredicto, str]:
    """Extrai veredicto e justificativa da resposta do Claude."""
    import re
    m_v = re.search(r'VEREDICTO:\s*(CORRETO|INCORRETO|PARCIAL)', texto, re.IGNORECASE)
    m_j = re.search(r'JUSTIFICATIVA:\s*(.+)', texto, re.IGNORECASE | re.DOTALL)

    if m_v:
        v_str = m_v.group(1).upper()
        veredicto = Veredicto(v_str)
    else:
        veredicto = Veredicto.PARCIAL  # conservador quando não consegue parsear

    justificativa = m_j.group(1).strip()[:400] if m_j else texto[:400]
    return veredicto, justificativa
