"""
Configuração central do verificador de citações.
Edite os caminhos abaixo conforme o local do pen drive no seu sistema.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ──────────────────────────────────────────────
# CAMINHOS — ajuste para o seu sistema
# ──────────────────────────────────────────────

# Windows: ex. "E:\\mestrado\\dissertacao final 1.docx"
# Linux/Mac: ex. "/media/usuario/PENDRIVE/mestrado/dissertacao final 1.docx"
DISSERTACAO_PATH = os.getenv(
    "DISSERTACAO_PATH",
    r"dissertacao final 1.docx"   # ← edite aqui ou use variável de ambiente
)

# Pasta com os arquivos das obras referenciadas
REFERENCIAS_FOLDER = os.getenv(
    "REFERENCIAS_FOLDER",
    r"fusion centers"              # ← edite aqui ou use variável de ambiente
)

# Pasta de saída para o relatório
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "relatorio_saida"))

# ──────────────────────────────────────────────
# API CLAUDE
# ──────────────────────────────────────────────
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# Modelo Claude a usar
CLAUDE_MODEL = "claude-sonnet-4-6"

# Máximo de tokens por chamada de verificação
MAX_TOKENS_VERIFICACAO = 1024

# Limite de caracteres do trecho de contexto enviado ao Claude
MAX_CHARS_CONTEXTO_DISSERTACAO = 1500
MAX_CHARS_CONTEXTO_FONTE = 4000

# ──────────────────────────────────────────────
# EXTENSÕES SUPORTADAS para a pasta de referências
# ──────────────────────────────────────────────
EXTENSOES_SUPORTADAS = {".pdf", ".docx", ".doc", ".txt"}
