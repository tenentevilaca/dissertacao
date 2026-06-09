"""
Extrai automaticamente os metadados da dissertação (autor, título, etc.)
para preencher os placeholders do template de análise.
"""

import json
import re
from pathlib import Path

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL
from modules.docx_utils import extrair_texto_docx

_MAX_CHARS = 8_000  # primeiros N chars são suficientes para metadados


def extrair_metadados(caminho_dissertacao: Path) -> dict:
    """
    Lê o início da dissertação e usa IA para identificar:
    autor, título, programa, instituição, orientador, área, ano, etc.
    """
    texto = extrair_texto_docx(caminho_dissertacao)[:_MAX_CHARS]

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""Você é um extrator de metadados acadêmicos. Leia o trecho abaixo do início de uma dissertação e extraia as informações solicitadas.

Se uma informação não estiver explicitamente no texto, use null.

Retorne SOMENTE um objeto JSON com as chaves abaixo, sem texto adicional:

{{
  "autor": "nome completo do autor",
  "titulo": "título completo da dissertação",
  "programa": "nome do programa/curso de pós-graduação",
  "instituicao": "nome da instituição",
  "orientador": "nome do orientador",
  "area": "área de concentração ou linha de pesquisa",
  "ano_inicio": "ano de início (se mencionado)",
  "ano_conclusao": "ano de conclusão ou defesa (se mencionado)",
  "palavras_chave": ["lista", "de", "palavras-chave"],
  "resumo_curto": "resumo do tema principal em 1-2 frases",
  "problema_pesquisa": "problema ou pergunta central da pesquisa, em 1-2 frases",
  "hipotese": "hipótese ou tese central, se identificável",
  "metodologia": "tipo de metodologia (survey, entrevistas, documental, mista, etc.)",
  "numero_capitulos_estimado": null
}}

TEXTO DA DISSERTAÇÃO (início):
{texto}
"""

    msg = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = msg.content[0].text.strip()
    # remove ```json ... ``` se presente
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"erro": "Não foi possível extrair metadados", "raw": raw}
