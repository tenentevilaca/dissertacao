"""
Verificador de Citações Acadêmicas
===================================
Uso:
    python main.py [--dissertacao CAMINHO] [--referencias PASTA] [--saida PASTA]

Variáveis de ambiente alternativas (ou arquivo .env):
    DISSERTACAO_PATH   — caminho para o .docx da dissertação
    REFERENCIAS_FOLDER — pasta com os arquivos das obras referenciadas
    OUTPUT_DIR         — pasta de saída do relatório
    ANTHROPIC_API_KEY  — chave da API Claude
"""

import argparse
import sys
from pathlib import Path

from colorama import init as colorama_init, Fore, Style

colorama_init(autoreset=True)


def _banner():
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════════════════════╗
║       VERIFICADOR DE CITAÇÕES ACADÊMICAS (ABNT)         ║
║       Powered by Claude AI                               ║
╚══════════════════════════════════════════════════════════╝
""")


def _parse_args():
    p = argparse.ArgumentParser(description="Verifica citações e referências em dissertações ABNT.")
    p.add_argument("--dissertacao",  help="Caminho do arquivo .docx da dissertação")
    p.add_argument("--referencias",  help="Caminho da pasta com as obras referenciadas")
    p.add_argument("--saida",        help="Pasta de saída para o relatório")
    p.add_argument("--sem-verificacao", action="store_true",
                   help="Pula a verificação semântica por IA (mais rápido, sem uso de API)")
    return p.parse_args()


def _configurar(args) -> None:
    """Sobrescreve config.py com os argumentos de linha de comando, se fornecidos."""
    import config
    if args.dissertacao:
        config.DISSERTACAO_PATH = args.dissertacao
    if args.referencias:
        config.REFERENCIAS_FOLDER = args.referencias
    if args.saida:
        config.OUTPUT_DIR = Path(args.saida)


def main():
    _banner()
    args = _parse_args()
    _configurar(args)

    import config
    from modules.dissertation_parser import DissertationParser
    from modules.reference_reader import ReferenceReader
    from modules.citation_verifier import CitationVerifier
    from modules.report_generator import ReportGenerator

    # ── Validações ───────────────────────────────────────────────────────
    dissertacao_path = Path(config.DISSERTACAO_PATH)
    referencias_path = Path(config.REFERENCIAS_FOLDER)

    if not dissertacao_path.exists():
        print(Fore.RED + f"[ERRO] Arquivo da dissertação não encontrado: {dissertacao_path}")
        print(Fore.YELLOW + "       Ajuste DISSERTACAO_PATH em config.py ou use --dissertacao")
        sys.exit(1)

    if not referencias_path.exists():
        print(Fore.RED + f"[ERRO] Pasta de referências não encontrada: {referencias_path}")
        print(Fore.YELLOW + "       Ajuste REFERENCIAS_FOLDER em config.py ou use --referencias")
        sys.exit(1)

    if not args.sem_verificacao and not config.ANTHROPIC_API_KEY:
        print(Fore.RED + "[ERRO] ANTHROPIC_API_KEY não configurada.")
        print(Fore.YELLOW + "       Defina a variável de ambiente ANTHROPIC_API_KEY ou crie um arquivo .env")
        print(Fore.YELLOW + "       Ou use --sem-verificacao para pular a etapa de IA.")
        sys.exit(1)

    # ── Etapa 1: Parsing da dissertação ─────────────────────────────────
    print(Fore.GREEN + "\n[1/4] Analisando a dissertação…")
    parser = DissertationParser(dissertacao_path)
    parser.carregar()

    citacoes   = parser.extrair_citacoes()
    referencias = parser.extrair_referencias()

    print(f"      → {len(citacoes)} citações encontradas no texto")
    print(f"      → {len(referencias)} entradas na lista de referências")

    # ── Etapa 2: Cruzamento citações × referências ───────────────────────
    print(Fore.GREEN + "\n[2/4] Cruzando citações com a lista de referências…")
    cruzamento = parser.cruzar_citacoes_referencias()

    citadas_sem_ref    = cruzamento["citadas_sem_referencia"]
    refs_sem_citacao   = cruzamento["referenciadas_sem_citacao"]
    pareamentos        = cruzamento["pareamentos"]

    print(f"      → {len(pareamentos)} pares (citação ↔ referência) encontrados")
    print(f"      → {len(citadas_sem_ref)} citações SEM entrada de referência")
    print(f"      → {len(refs_sem_citacao)} referências SEM citação no texto")

    # ── Etapa 3: Leitura dos documentos-fonte ───────────────────────────
    print(Fore.GREEN + "\n[3/4] Lendo os documentos da pasta de referências…")
    reader = ReferenceReader(referencias_path)
    reader.carregar_todos(verbose=True)

    # ── Etapa 4: Verificação semântica por IA ────────────────────────────
    verificacoes = []
    if not args.sem_verificacao:
        print(Fore.GREEN + f"\n[4/4] Verificando {len(pareamentos)} citações com Claude AI…")
        verifier = CitationVerifier(reader)
        verificacoes = verifier.verificar_lote(pareamentos, verbose=True)
    else:
        print(Fore.YELLOW + "\n[4/4] Verificação semântica pulada (--sem-verificacao).")

    # ── Geração do relatório ─────────────────────────────────────────────
    print(Fore.GREEN + "\n[✓] Gerando relatórios…")
    output_dir = Path(config.OUTPUT_DIR)
    gen = ReportGenerator(
        citadas_sem_ref=citadas_sem_ref,
        refs_sem_citacao=refs_sem_citacao,
        verificacoes=verificacoes,
        caminho_dissertacao=str(dissertacao_path),
    )

    path_txt  = gen.gerar_txt(output_dir / "relatorio.txt")
    path_html = gen.gerar_html(output_dir / "relatorio.html")
    path_json = gen.gerar_json(output_dir / "relatorio.json")

    print(Fore.CYAN + f"""
Relatórios salvos em: {output_dir}/
  • relatorio.txt   — texto simples
  • relatorio.html  — relatório visual (abra no navegador)
  • relatorio.json  — dados estruturados para auditoria
""")

    # ── Resumo no terminal ────────────────────────────────────────────────
    corretas   = sum(1 for v in verificacoes if v.veredicto.value == "CORRETO")
    incorretas = sum(1 for v in verificacoes if v.veredicto.value == "INCORRETO")
    parciais   = sum(1 for v in verificacoes if v.veredicto.value == "PARCIAL")
    sem_fonte  = sum(1 for v in verificacoes if v.veredicto.value == "SEM_FONTE")

    print(Style.BRIGHT + "RESUMO:")
    print(f"  ✓ Corretas          : {Fore.GREEN}{corretas}")
    print(f"  ⚠ Parciais          : {Fore.YELLOW}{parciais}")
    print(f"  ✗ Incorretas        : {Fore.RED}{incorretas}")
    print(f"  ? Sem arquivo-fonte : {Fore.WHITE}{sem_fonte}")
    print(f"  ↑ Sem referência    : {Fore.CYAN}{len(citadas_sem_ref)}")
    print(f"  ↓ Ref. sem citação  : {Fore.CYAN}{len(refs_sem_citacao)}")


if __name__ == "__main__":
    main()
