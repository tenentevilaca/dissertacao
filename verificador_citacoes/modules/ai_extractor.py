"""
Extração de citações e referências usando Claude AI.
Usa leitura XML direta do .docx (mais robusta que python-docx).
"""

import json
import re
import zipfile
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import anthropic

from modules.dissertation_parser import CitacaoInText, EntradaReferencia

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MAX_CHARS = 150_000

_PROMPT = """Você é um especialista em análise de textos acadêmicos brasileiros.

Analise o texto abaixo de uma dissertação/tese acadêmica e realize DUAS tarefas:

TAREFA 1 — Citações no corpo do texto
Identifique TODAS as referências bibliográficas no corpo do texto.
Inclua qualquer formato: (AUTOR, ano), Autor (ano), notas de rodapé [1], citações entre aspas com fonte, etc.

TAREFA 2 — Lista de referências
Extraia CADA ENTRADA da lista de referências (seção "REFERÊNCIAS", "REFERÊNCIAS BIBLIOGRÁFICAS" ou similar no final).

Retorne SOMENTE um JSON, sem texto antes ou depois:
{
  "citacoes": [
    {
      "texto_original": "exatamente como aparece no texto, ex: (CASTELLS, 2001, p. 117)",
      "autores": ["CASTELLS"],
      "ano": "2001",
      "pagina": "117",
      "contexto": "trecho de 200-300 caracteres ao redor da citação no texto"
    }
  ],
  "referencias": [
    {
      "texto_completo": "entrada completa como aparece na lista de referências",
      "autor_sobrenome": "CASTELLS",
      "ano": "2001",
      "titulo": "título identificado da obra"
    }
  ]
}

TEXTO DA DISSERTAÇÃO:
"""


def _texto_de_xml(xml_bytes: bytes) -> str:
    """Extrai todos os textos de um XML do Word."""
    try:
        root = ET.fromstring(xml_bytes)
        partes = []
        for elem in root.iter(f"{{{_NS}}}t"):
            if elem.text:
                partes.append(elem.text)
            # Preserva quebras de parágrafo
        for elem in root.iter(f"{{{_NS}}}p"):
            partes.append("\n")
        return "".join(partes)
    except Exception:
        return ""


def extrair_texto_docx(caminho: str | Path) -> str:
    """
    Extrai texto completo de um .docx lendo o XML diretamente.
    Captura: corpo do documento, tabelas, cabeçalhos, rodapés e notas de rodapé.
    """
    partes: list[str] = []

    try:
        with zipfile.ZipFile(str(caminho), "r") as z:
            nomes = z.namelist()

            # Documento principal
            if "word/document.xml" in nomes:
                partes.append(_texto_de_xml(z.read("word/document.xml")))

            # Notas de rodapé
            if "word/footnotes.xml" in nomes:
                partes.append("\n--- NOTAS DE RODAPÉ ---\n")
                partes.append(_texto_de_xml(z.read("word/footnotes.xml")))

            # Notas finais
            if "word/endnotes.xml" in nomes:
                partes.append("\n--- NOTAS FINAIS ---\n")
                partes.append(_texto_de_xml(z.read("word/endnotes.xml")))

            # Cabeçalhos e rodapés
            for nome in nomes:
                if nome.startswith("word/header") or nome.startswith("word/footer"):
                    partes.append(_texto_de_xml(z.read(nome)))

    except zipfile.BadZipFile:
        # Tenta python-docx como fallback
        try:
            import docx as docx_lib
            doc = docx_lib.Document(str(caminho))
            partes.append("\n".join(p.text for p in doc.paragraphs))
        except Exception as e:
            partes.append(f"[ERRO AO LER ARQUIVO: {e}]")

    texto = "\n".join(partes)
    # Remove linhas vazias excessivas
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


class AIExtractor:
    """Usa Claude para extrair citações e referências de qualquer formato acadêmico."""

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def extrair(self, texto: str, log_fn=None) -> dict:
        if not texto.strip():
            if log_fn:
                log_fn("   ✗ Texto vazio — não foi possível ler o arquivo", "warn")
            return {"citacoes": [], "referencias": []}

        if log_fn:
            log_fn(f"   Enviando {len(texto):,} caracteres ao Claude…")

        try:
            resposta = self.client.messages.create(
                model=self.model,
                max_tokens=8192,
                messages=[{"role": "user", "content": _PROMPT + texto[:_MAX_CHARS]}],
            )
        except Exception as e:
            if log_fn:
                log_fn(f"   ✗ Erro na API Claude: {e}")
            raise

        raw = resposta.content[0].text.strip()

        if log_fn:
            log_fn(f"   Claude respondeu ({len(raw)} chars)")
            log_fn(f"   Prévia: {raw[:150].replace(chr(10),' ')}")

        # Limpa markdown se presente
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.MULTILINE)
        raw = re.sub(r"\s*```\s*$", "", raw, flags=re.MULTILINE)

        # Tenta parsear JSON
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\{[\s\S]*\}", raw)
            if m:
                try:
                    return json.loads(m.group())
                except json.JSONDecodeError:
                    pass
            if log_fn:
                log_fn(f"   ✗ JSON inválido. Resposta bruta: {raw[:400]}")
            return {"citacoes": [], "referencias": []}

    def ler_docx(self, caminho: str | Path) -> str:
        """Retorna o texto extraído do .docx (para uso externo antes de extrair)."""
        return extrair_texto_docx(caminho)

    def extrair_de_docx(self, caminho: str | Path, log_fn=None) -> dict:
        texto = extrair_texto_docx(caminho)
        if log_fn:
            log_fn(f"   Texto extraído do .docx: {len(texto):,} caracteres")
            if len(texto) < 200:
                log_fn(f"   ⚠ Texto muito curto: {repr(texto[:200])}")
        return self.extrair(texto, log_fn=log_fn)

    def para_objetos(self, dados: dict) -> tuple[list[CitacaoInText], list[EntradaReferencia]]:
        citacoes: list[CitacaoInText] = []
        for i, c in enumerate(dados.get("citacoes", [])):
            try:
                autores = c.get("autores") or ["DESCONHECIDO"]
                if isinstance(autores, str):
                    autores = [autores]
                citacoes.append(CitacaoInText(
                    texto_original=str(c.get("texto_original", "")),
                    autores=[str(a).upper() for a in autores],
                    ano=str(c.get("ano", "s.d.")),
                    pagina=c.get("pagina") or None,
                    contexto=str(c.get("contexto", "")),
                    paragrafo_idx=i,
                ))
            except Exception:
                continue

        referencias: list[EntradaReferencia] = []
        for r in dados.get("referencias", []):
            try:
                referencias.append(EntradaReferencia(
                    texto_completo=str(r.get("texto_completo", "")),
                    autor_sobrenome=str(r.get("autor_sobrenome", "DESCONHECIDO")).upper(),
                    ano=str(r.get("ano", "s.d.")),
                    titulo=str(r.get("titulo", "")),
                ))
            except Exception:
                continue

        return citacoes, referencias

    def cruzar(self, citacoes: list[CitacaoInText], referencias: list[EntradaReferencia]) -> dict:
        idx_refs: dict[str, EntradaReferencia] = {}
        for ref in referencias:
            chave = f"{ref.autor_sobrenome.upper()}_{ref.ano}"
            idx_refs[chave] = ref

        chaves_citadas: set[str] = set()
        citadas_sem_ref: list[CitacaoInText] = []
        pareamentos: list[tuple] = []

        for cit in citacoes:
            encontrou = False
            for autor in cit.autores:
                chave = f"{autor.upper()}_{cit.ano}"
                chaves_citadas.add(chave)
                if chave in idx_refs:
                    pareamentos.append((cit, idx_refs[chave]))
                    encontrou = True
                    break
            if not encontrou:
                citadas_sem_ref.append(cit)

        refs_sem_citacao = [r for k, r in idx_refs.items() if k not in chaves_citadas]

        return {
            "citadas_sem_referencia": citadas_sem_ref,
            "referenciadas_sem_citacao": refs_sem_citacao,
            "pareamentos": pareamentos,
        }
