"""Configuração central do Analisador de Dissertação."""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = "claude-sonnet-4-6"

# Limite de caracteres por capítulo enviado à IA (evita estourar contexto)
MAX_CHARS_CAPITULO = 55_000

# Limite de chars por trecho de material de apoio
MAX_CHARS_APOIO = 3_500

# Número de trechos de apoio por afirmação buscada
TOP_K_APOIO = 3

OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "saida_analise"))
