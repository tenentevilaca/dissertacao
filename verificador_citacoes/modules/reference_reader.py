"""
Lê os documentos da pasta de referências (PDFs, docx, txt) e indexa
seu conteúdo para ser consultado durante a verificação de citações.
"""

import re
from dataclasses import dataclass, field
from pathlib import Path

import fitz          # PyMuPDF
import docx as docx_lib

from config import EXTENSOES_SUPORTADAS


@dataclass
class DocumentoReferencia:
    """Texto extraído de um arquivo de referência."""
    caminho: Path
    nome_arquivo: str
    texto_completo: str
    paginas: dict[int, str] = field(default_factory=dict)   # page_num → texto
    metadados: dict = field(default_factory=dict)


class ReferenceReader:
    def __init__(self, pasta: str | Path):
        self.pasta = Path(pasta)
        self.documentos: list[DocumentoReferencia] = []
        # Índice: chave de busca (ex. sobrenome+ano) → lista de documentos candidatos
        self._indice: dict[str, list[DocumentoReferencia]] = {}

    # ── Leitura de arquivos ──────────────────

    def carregar_todos(self, verbose: bool = True) -> "ReferenceReader":
        if not self.pasta.exists():
            raise FileNotFoundError(f"Pasta não encontrada: {self.pasta}")

        arquivos = [
            f for f in self.pasta.rglob("*")
            if f.is_file() and f.suffix.lower() in EXTENSOES_SUPORTADAS
        ]

        if verbose:
            print(f"[INFO] {len(arquivos)} arquivo(s) encontrado(s) em '{self.pasta}'")

        for arq in arquivos:
            doc = self._ler_arquivo(arq)
            if doc:
                self.documentos.append(doc)
                if verbose:
                    print(f"  ✓ {arq.name} ({len(doc.texto_completo):,} chars)")

        self._construir_indice()
        return self

    def _ler_arquivo(self, caminho: Path) -> DocumentoReferencia | None:
        ext = caminho.suffix.lower()
        try:
            if ext == ".pdf":
                return self._ler_pdf(caminho)
            elif ext in {".docx", ".doc"}:
                return self._ler_docx(caminho)
            elif ext == ".txt":
                return self._ler_txt(caminho)
        except Exception as e:
            print(f"  [ERRO] {caminho.name}: {e}")
        return None

    def _ler_pdf(self, caminho: Path) -> DocumentoReferencia:
        paginas: dict[int, str] = {}
        doc_fitz = fitz.open(str(caminho))
        for i, page in enumerate(doc_fitz, start=1):
            paginas[i] = page.get_text()
        doc_fitz.close()
        texto_completo = "\n".join(paginas.values())
        return DocumentoReferencia(
            caminho=caminho,
            nome_arquivo=caminho.name,
            texto_completo=texto_completo,
            paginas=paginas,
        )

    def _ler_docx(self, caminho: Path) -> DocumentoReferencia:
        doc = docx_lib.Document(str(caminho))
        texto = "\n".join(p.text for p in doc.paragraphs)
        return DocumentoReferencia(
            caminho=caminho,
            nome_arquivo=caminho.name,
            texto_completo=texto,
        )

    def _ler_txt(self, caminho: Path) -> DocumentoReferencia:
        texto = caminho.read_text(encoding="utf-8", errors="replace")
        return DocumentoReferencia(
            caminho=caminho,
            nome_arquivo=caminho.name,
            texto_completo=texto,
        )

    # ── Indexação ────────────────────────────

    def _construir_indice(self):
        """
        Cria um índice de busca baseado em tokens do nome do arquivo.
        Ex.: "SILVA_2020_titulo.pdf" → chaves "silva", "2020"
        """
        self._indice = {}
        for doc in self.documentos:
            tokens = re.split(r'[\s_\-\.\,\(\)]+', doc.nome_arquivo.lower())
            for tok in tokens:
                tok = tok.strip()
                if tok:
                    self._indice.setdefault(tok, []).append(doc)

    # ── Busca ────────────────────────────────

    def buscar_documento(
        self,
        sobrenome: str,
        ano: str,
        titulo_palavras: list[str] | None = None,
    ) -> list[DocumentoReferencia]:
        """
        Busca documentos que provavelmente correspondem à referência.
        Retorna uma lista ordenada por relevância (mais relevante primeiro).
        """
        candidatos: dict[int, DocumentoReferencia] = {}   # id(doc) → doc
        pontuacao: dict[int, int] = {}

        def _adicionar(doc: DocumentoReferencia, pts: int):
            key = id(doc)
            candidatos[key] = doc
            pontuacao[key] = pontuacao.get(key, 0) + pts

        sob_lower = sobrenome.lower()
        # Correspondência por sobrenome no nome do arquivo
        for tok, docs in self._indice.items():
            if sob_lower in tok or tok in sob_lower:
                for d in docs:
                    _adicionar(d, 3)

        # Correspondência por ano
        for d in self._indice.get(ano, []):
            _adicionar(d, 2)

        # Correspondência por palavras do título
        if titulo_palavras:
            for palavra in titulo_palavras:
                p_lower = palavra.lower()
                if len(p_lower) < 4:
                    continue
                for tok, docs in self._indice.items():
                    if p_lower in tok:
                        for d in docs:
                            _adicionar(d, 1)

        if not candidatos:
            return []

        ordenados = sorted(candidatos.values(), key=lambda d: pontuacao[id(d)], reverse=True)
        return ordenados

    def trecho_por_pagina(self, doc: DocumentoReferencia, pagina: str | None, janela: int = 4000) -> str:
        """
        Retorna o texto de uma página específica (ou o início do documento).
        """
        if pagina and doc.paginas:
            try:
                num = int(re.sub(r'\D', '', pagina))
                texto_pag = doc.paginas.get(num, "")
                if texto_pag:
                    return texto_pag[:janela]
            except ValueError:
                pass
        # Sem página específica: retorna os primeiros `janela` chars
        return doc.texto_completo[:janela]
