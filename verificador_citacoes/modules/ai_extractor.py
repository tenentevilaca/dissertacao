"""
Extrai citações e referências usando Claude AI — sem regex, sem regras fixas.
Claude lê o texto e identifica tudo independente do formato de citação usado.
"""

import json
import re
from pathlib import Path
from typing import Optional

import anthropic
import docx as docx_lib

from modules.dissertation_parser import CitacaoInText, EntradaReferencia


_PROMPT = """Você é especialista em análise de textos acadêmicos brasileiros.

Analise o texto abaixo de uma dissertação/tese e realize DUAS tarefas:

TAREFA 1 — Citações no corpo do texto
Identifique TODAS as referências a obras de terceiros no corpo do texto (antes da seção de referências). Inclua qualquer formato encontrado: (AUTOR, ano), Autor (ano), notas numeradas, citações diretas com aspas, etc.

TAREFA 2 — Lista de referências bibliográficas
Extraia CADA ENTRADA da seção "REFERÊNCIAS" ou "REFERÊNCIAS BIBLIOGRÁFICAS" que aparecer no final do documento.

Retorne SOMENTE um JSON válido, sem nenhum texto antes ou depois:
{
  "citacoes": [
    {
      "texto_original": "texto exato como aparece no documento, ex: (SILVA, 2020, p. 45)",
      "autores": ["SILVA"],
      "ano": "2020",
      "pagina": "45",
      "contexto": "trecho de ~300 caracteres do texto ao redor da citação"
    }
  ],
  "referencias": [
    {
      "texto_completo": "entrada completa da referência como aparece no documento",
      "autor_sobrenome": "SILVA",
      "ano": "2020",
      "titulo": "Título da obra identificado"
    }
  ]
}

Se não encontrar citações ou referências, retorne listas vazias — nunca omita as chaves.

TEXTO DA DISSERTAÇÃO:
"""

_MAX_CHARS = 160_000   # margem segura para a janela de 200K tokens do Claude


class AIExtractor:
    """Usa Claude para extrair citações e referências de qualquer formato acadêmico."""

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    # ── Leitura ──────────────────────────────────────────────────────────

    def ler_docx(self, caminho: str | Path) -> str:
        """
        Extrai texto completo de um .docx incluindo:
        parágrafos, tabelas, cabeçalhos, rodapés e notas de rodapé.
        """
        doc = docx_lib.Document(str(caminho))
        partes: list[str] = []

        # Parágrafos normais
        for p in doc.paragraphs:
            if p.text.strip():
                partes.append(p.text)

        # Tabelas (células)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for p in celula.paragraphs:
                        if p.text.strip():
                            partes.append(p.text)

        # Cabeçalhos e rodapés de cada seção
        for secao in doc.sections:
            for p in secao.header.paragraphs:
                if p.text.strip():
                    partes.append(p.text)
            for p in secao.footer.paragraphs:
                if p.text.strip():
                    partes.append(p.text)

        # Notas de rodapé (via XML direto)
        try:
            from docx.oxml.ns import qn
            for fn in doc.element.findall(
                f'.//{qn("w:footnote")}', doc.element.nsmap
            ):
                texto_fn = "".join(t.text or "" for t in fn.iter(qn("w:t")))
                if texto_fn.strip():
                    partes.append(texto_fn)
        except Exception:
            pass

        return "\n".join(partes)

    # ── Extração via Claude ───────────────────────────────────────────────

    def extrair(self, texto: str, log_fn=None) -> dict:
        """
        Envia o texto ao Claude e obtém citações + referências como JSON.
        log_fn: função opcional para registrar progresso (str → None).
        """
        if log_fn:
            chars = len(texto)
            log_fn(f"   Texto da dissertação: {chars:,} caracteres")
            if chars > _MAX_CHARS:
                log_fn(f"   Texto muito longo — enviando os primeiros {_MAX_CHARS:,} caracteres")

        if log_fn:
            log_fn("   Enviando texto ao Claude para análise inteligente…")

        try:
            resposta = self.client.messages.create(
                model=self.model,
                max_tokens=8192,
                messages=[{
                    "role": "user",
                    "content": _PROMPT + texto[:_MAX_CHARS]
                }]
            )
        except Exception as e:
            if log_fn:
                log_fn(f"   ✗ Erro na chamada à API Claude: {e}")
            raise

        raw = resposta.content[0].text.strip()

        if log_fn:
            log_fn(f"   Claude respondeu ({len(raw)} chars). Processando JSON…")
            # Mostra prévia da resposta para diagnóstico
            log_fn(f"   Prévia: {raw[:200].replace(chr(10), ' ')}")

        # Remove blocos markdown se o modelo os adicionar
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        try:
            dados = json.loads(raw)
        except json.JSONDecodeError:
            # Tenta extrair JSON embutido no texto
            m = re.search(r'\{[\s\S]*\}', raw)
            if m:
                try:
                    dados = json.loads(m.group())
                except json.JSONDecodeError:
                    if log_fn:
                        log_fn(f"   ✗ Falha ao interpretar resposta do Claude. Resposta raw: {raw[:300]}")
                    dados = {"citacoes": [], "referencias": []}
            else:
                if log_fn:
                    log_fn(f"   ✗ Claude não retornou JSON válido. Resposta: {raw[:300]}")
                dados = {"citacoes": [], "referencias": []}

        return dados

    def extrair_de_docx(self, caminho: str | Path, log_fn=None) -> dict:
        texto = self.ler_docx(caminho)
        return self.extrair(texto, log_fn=log_fn)

    # ── Conversão para objetos ────────────────────────────────────────────

    def para_objetos(
        self, dados: dict
    ) -> tuple[list[CitacaoInText], list[EntradaReferencia]]:
        """Converte o JSON do Claude nos objetos usados pelo resto do sistema."""
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

    # ── Cruzamento citações × referências ────────────────────────────────

    def cruzar(
        self,
        citacoes: list[CitacaoInText],
        referencias: list[EntradaReferencia],
    ) -> dict:
        """Cruza citações com referências por autor + ano."""
        idx_refs: dict[str, EntradaReferencia] = {}
        for ref in referencias:
            chave = f"{ref.autor_sobrenome.upper()}_{ref.ano}"
            idx_refs[chave] = ref

        chaves_citadas: set[str] = set()
        citadas_sem_ref: list[CitacaoInText] = []
        pareamentos: list[tuple[CitacaoInText, EntradaReferencia]] = []

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

        refs_sem_citacao = [
            ref for chave, ref in idx_refs.items()
            if chave not in chaves_citadas
        ]

        return {
            "citadas_sem_referencia": citadas_sem_ref,
            "referenciadas_sem_citacao": refs_sem_citacao,
            "pareamentos": pareamentos,
        }
