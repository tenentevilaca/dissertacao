"""Utilitários para extração de texto e estrutura de arquivos .docx e .pdf."""

import zipfile
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import fitz  # PyMuPDF

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_RST = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def extrair_texto_docx(caminho: Path) -> str:
    """Extrai texto completo de .docx lendo XML diretamente."""
    partes: list[str] = []
    try:
        with zipfile.ZipFile(str(caminho), "r") as z:
            for nome in z.namelist():
                if nome.startswith("word/") and nome.endswith(".xml"):
                    try:
                        xml = z.read(nome)
                        root = ET.fromstring(xml)
                        for elem in root.iter(f"{{{_NS}}}p"):
                            linha = "".join(
                                t.text for t in elem.iter(f"{{{_NS}}}t") if t.text
                            )
                            if linha.strip():
                                partes.append(linha)
                    except Exception:
                        pass
    except Exception:
        pass
    return "\n".join(partes)


def extrair_paragrafos_docx(caminho: Path) -> list[dict]:
    """
    Retorna lista de dicts com {texto, estilo, nivel} para cada parágrafo.
    Útil para detectar títulos de capítulos por estilo.
    """
    try:
        import docx as _docx
        doc = _docx.Document(str(caminho))
        resultado = []
        for para in doc.paragraphs:
            texto = para.text.strip()
            if not texto:
                continue
            estilo = para.style.name if para.style else ""
            nivel = 0
            if "Heading" in estilo or "Título" in estilo or "Title" in estilo:
                try:
                    nivel = int("".join(c for c in estilo if c.isdigit()) or "1")
                except ValueError:
                    nivel = 1
            resultado.append({"texto": texto, "estilo": estilo, "nivel": nivel})
        return resultado
    except Exception:
        # fallback: texto puro sem estilos
        texto = extrair_texto_docx(caminho)
        return [{"texto": l, "estilo": "", "nivel": 0} for l in texto.splitlines() if l.strip()]


def extrair_texto_pdf(caminho: Path) -> str:
    """Extrai texto de arquivo PDF."""
    try:
        doc = fitz.open(str(caminho))
        partes = []
        for page in doc:
            partes.append(page.get_text())
        return "\n".join(partes)
    except Exception:
        return ""


def extrair_texto_arquivo(caminho: Path) -> str:
    """Extrai texto de .docx, .pdf ou .txt."""
    ext = caminho.suffix.lower()
    if ext == ".docx":
        return extrair_texto_docx(caminho)
    if ext == ".pdf":
        return extrair_texto_pdf(caminho)
    if ext in {".txt", ".md"}:
        try:
            return caminho.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return ""
    return ""
