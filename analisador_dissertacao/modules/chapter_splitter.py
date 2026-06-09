"""
Divide a dissertação em capítulos, usando estilos de heading do .docx.
Fallback: heurística textual (linhas em maiúsculas ou numeradas).
"""

import re
from pathlib import Path

import docx as _docx

from modules.docx_utils import extrair_paragrafos_docx


def _e_titulo_capitulo(para: dict) -> bool:
    """Retorna True se o parágrafo parece ser título de capítulo (nível 1)."""
    nivel = para.get("nivel", 0)
    if nivel == 1:
        return True

    # Fallback heurístico para documentos sem estilos
    texto = para["texto"].strip()
    if not texto or len(texto) > 120:
        return False

    # "CAPÍTULO 1", "1 INTRODUÇÃO", "1. INTRODUÇÃO", "INTRODUÇÃO"
    if re.match(r"^(cap[íi]tulo\s+\d+|cap\.\s*\d+)", texto, re.IGNORECASE):
        return True
    if re.match(r"^\d+[\.\s]+[A-ZÁÉÍÓÚ]", texto):
        return True
    if texto.isupper() and 4 < len(texto) < 80:
        return True

    return False


def dividir_em_capitulos(caminho: Path) -> list[dict]:
    """
    Retorna lista de capítulos:
      [{"titulo": str, "numero": int, "paragrafos": [str], "texto": str}]
    """
    paragrafos = extrair_paragrafos_docx(caminho)

    capitulos: list[dict] = []
    capitulo_atual: dict | None = None
    numero = 0

    for para in paragrafos:
        if _e_titulo_capitulo(para):
            if capitulo_atual is not None:
                capitulo_atual["texto"] = "\n\n".join(capitulo_atual["paragrafos"])
                capitulos.append(capitulo_atual)
            numero += 1
            capitulo_atual = {
                "titulo": para["texto"],
                "numero": numero,
                "paragrafos": [],
                "texto": "",
            }
        else:
            if capitulo_atual is None:
                # texto antes do primeiro capítulo → pré-textual
                capitulo_atual = {
                    "titulo": "Elementos Pré-textuais",
                    "numero": 0,
                    "paragrafos": [],
                    "texto": "",
                }
            capitulo_atual["paragrafos"].append(para["texto"])

    if capitulo_atual is not None:
        capitulo_atual["texto"] = "\n\n".join(capitulo_atual["paragrafos"])
        capitulos.append(capitulo_atual)

    return capitulos
