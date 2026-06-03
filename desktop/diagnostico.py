#!/usr/bin/env python3
"""
DIAGNÓSTICO DO VERIFICADOR DE CITAÇÕES
Execute: python diagnostico.py "caminho/para/dissertacao.docx"
Depois copie e cole toda a saída aqui no chat.
"""

import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

VERSAO_ESPERADA = "2025-06-02-v8.6"

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def xml_para_texto(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
        linhas = []
        for para in root.iter(f"{{{_NS}}}p"):
            tokens = [t.text for t in para.iter(f"{{{_NS}}}t") if t.text]
            linhas.append("".join(tokens))
        return "\n".join(linhas)
    except Exception as e:
        return f"[ERRO XML: {e}]"

def ler_docx(caminho):
    try:
        with zipfile.ZipFile(str(caminho)) as z:
            nomes = z.namelist()
            partes = []
            if "word/document.xml" in nomes:
                partes.append(xml_para_texto(z.read("word/document.xml")))
            if "word/footnotes.xml" in nomes:
                partes.append(xml_para_texto(z.read("word/footnotes.xml")))
            return "\n".join(partes)
    except Exception as e:
        return f"[ERRO: {e}]"


def main():
    print("=" * 70)
    print("DIAGNÓSTICO DO VERIFICADOR DE CITAÇÕES")
    print("=" * 70)

    # ── Verifica versão do analisar.py ──────────────────────────────────
    try:
        import analisar
        versao_real = "DESCONHECIDA"
        src = Path(analisar.__file__).read_text(encoding="utf-8", errors="replace")
        m = re.search(r"versão ([\w\-]+)", src)
        if m:
            versao_real = m.group(1)
        print(f"\n[1] Versão do analisar.py carregado: {versao_real}")
        if versao_real == VERSAO_ESPERADA:
            print(f"    ✓ Versão correta!")
        else:
            print(f"    ✗ VERSÃO INCORRETA! Esperada: {VERSAO_ESPERADA}")
            print(f"    → Baixe o arquivo mais recente do GitHub e substitua.")
    except Exception as e:
        print(f"    ✗ Não encontrou analisar.py: {e}")

    # ── Verifica arquivo DOCX ────────────────────────────────────────────
    if len(sys.argv) < 2:
        print("\n[2] Nenhum arquivo informado.")
        print("    Uso: python diagnostico.py \"caminho/dissertacao.docx\"")
        return

    caminho = Path(sys.argv[1])
    print(f"\n[2] Arquivo: {caminho}")
    if not caminho.exists():
        print("    ✗ ARQUIVO NÃO ENCONTRADO!")
        return
    print(f"    Tamanho: {caminho.stat().st_size // 1024} KB")
    print(f"    Extensão: {caminho.suffix}")

    # ── Lê o DOCX ───────────────────────────────────────────────────────
    print("\n[3] Lendo o DOCX...")
    texto = ler_docx(caminho)
    print(f"    Caracteres extraídos: {len(texto):,}")

    if len(texto) < 500:
        print("    ✗ TEXTO MUITO CURTO! Primeiros 200 chars:")
        print(f"    {repr(texto[:200])}")
        return

    paragrafos = texto.split("\n")
    nao_vazios = [p for p in paragrafos if p.strip()]
    print(f"    Total parágrafos: {len(paragrafos)}")
    print(f"    Parágrafos não-vazios: {len(nao_vazios)}")

    # ── Primeiros parágrafos ─────────────────────────────────────────────
    print("\n[4] Primeiros 15 parágrafos não-vazios:")
    for i, p in enumerate(nao_vazios[:15], 1):
        print(f"    {i:>3}. {p[:100]}")

    # ── Verifica caracteres especiais ────────────────────────────────────
    print("\n[5] Verificando parênteses no texto:")
    paren_normal = sum(1 for p in paragrafos if "(" in p and ")" in p)
    paren_wide   = sum(1 for p in paragrafos if "（" in p or "）" in p)  # fullwidth
    print(f"    Com parênteses normais  ( ): {paren_normal}")
    print(f"    Com parênteses largos ｀）: {paren_wide}")
    if paren_wide and not paren_normal:
        print("    ✗ ATENÇÃO: Documento usa parênteses especiais (fullwidth)!")
        print("      O verificador não reconhece esses caracteres.")

    # ── Busca REFERÊNCIAS ─────────────────────────────────────────────────
    print("\n[6] Buscando seção REFERÊNCIAS:")
    RE_REFS = re.compile(r"^\s*REFER[EÊ]NCIAS\s*(BIBLIOGR[ÁA]FICAS?)?\s*$", re.IGNORECASE)
    ocorrencias = [(i, p) for i, p in enumerate(paragrafos) if RE_REFS.match(p.strip())]
    if not ocorrencias:
        print("    ✗ Seção REFERÊNCIAS NÃO ENCONTRADA no documento!")
        print("    → Primeiros parágrafos que contêm a palavra 'referência':")
        for i, p in enumerate(paragrafos):
            if "referênc" in p.lower() or "referenc" in p.lower():
                print(f"      Linha {i}: {p[:100]}")
    else:
        total = len(paragrafos)
        print(f"    Encontrada {len(ocorrencias)} vez(es):")
        for i, p in ocorrencias:
            print(f"      Linha {i} de {total} ({100*i//total}%): {p.strip()!r}")

        metade = total // 2
        tardios = [i for i, _ in ocorrencias if i >= metade]
        idx_refs = tardios[-1] if tardios else ocorrencias[-1][0]
        limite = idx_refs
        print(f"    → idx_refs escolhido: {idx_refs}")
        print(f"    → Parágrafos varridos para citações: 0 até {limite-1}")

        # ── Testa regex nas primeiras linhas ─────────────────────────────
        RE_PAREN = re.compile(
            r"\(([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,}"
            r"(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,})?"
            r"(?:,?\s*et\s+al\.?)?"
            r"(?:\s*;\s*[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,})*)"
            r"[,;]\s*(\d{4}[a-z]?)"
            r"(?:[,;:\s]+(?:p\.?p?|pág\.?|f\.?)\s*([\d\-–]+))?\)"
        )
        RE_NARR = re.compile(
            r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,}"
            r"(?:\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,})?"
            r"(?:,?\s*et\s+al\.?)?)"
            r"\s+\((\d{4}[a-z]?)"
        )

        print(f"\n[7] Testando regex nos primeiros {min(limite,200)} parágrafos:")
        achou_paren = 0
        achou_narr = 0
        amostras = []
        for p in paragrafos[:limite]:
            if RE_PAREN.search(p):
                achou_paren += 1
                if len(amostras) < 5:
                    amostras.append(("PAREN", p.strip()[:100]))
            if RE_NARR.search(p):
                achou_narr += 1
                if len(amostras) < 5:
                    amostras.append(("NARR", p.strip()[:100]))

        print(f"    Parágrafos com citação parentética (AUTOR, ANO): {achou_paren}")
        print(f"    Parágrafos com citação narrativa Autor (ANO):    {achou_narr}")

        if achou_paren == 0 and achou_narr == 0:
            print("\n    ✗ NENHUMA CITAÇÃO DETECTADA nos parágrafos varridos.")
            print("    → Parágrafos com parênteses nessa faixa:")
            conta = 0
            for p in paragrafos[:limite]:
                if ("(" in p or "（" in p) and p.strip() and conta < 10:
                    print(f"      {p.strip()[:120]}")
                    conta += 1
            if conta == 0:
                print("      (nenhum parágrafo com parênteses encontrado)")
        else:
            print("\n    ✓ Amostras encontradas:")
            for tipo, txt in amostras:
                print(f"      [{tipo}] {txt}")

    print("\n" + "=" * 70)
    print("Cole toda esta saída no chat para diagnóstico.")
    print("=" * 70)


if __name__ == "__main__":
    main()
