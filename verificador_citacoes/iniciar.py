"""
Lançador automático — instala dependências e abre o app no navegador.
Execute diretamente: python iniciar.py
Ou clique duas vezes em INICIAR.bat (Windows) / INICIAR.command (Mac).
"""

import os
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

PORTA = 8000
PASTA = Path(__file__).parent.resolve()

# ── Cabeçalho ────────────────────────────────────────────────────────────
print("""
╔══════════════════════════════════════════════════════════╗
║       VERIFICADOR DE CITAÇÕES ACADÊMICAS                ║
╚══════════════════════════════════════════════════════════╝
""")

# ── Verificação do Python ─────────────────────────────────────────────────
if sys.version_info < (3, 11):
    print(f"[ERRO] Python 3.11+ é necessário. Versão atual: {sys.version}")
    print("       Baixe em: https://www.python.org/downloads/")
    input("\nPressione Enter para fechar...")
    sys.exit(1)

print(f"Python {sys.version.split()[0]} detectado. ✓")

# ── Instalação automática de dependências ─────────────────────────────────
req = PASTA / "requirements.txt"
print("\nVerificando dependências (pode demorar na primeira vez)...")

resultado = subprocess.run(
    [sys.executable, "-m", "pip", "install", "-q", "-r", str(req)],
    capture_output=True, text=True
)
if resultado.returncode != 0:
    print("[AVISO] Problema ao instalar dependências:")
    print(resultado.stderr[:500])
    print("Tentando continuar mesmo assim...")
else:
    print("Dependências prontas. ✓")

# ── Abre o navegador após 2s ──────────────────────────────────────────────
def _abrir_browser():
    time.sleep(2)
    url = f"http://localhost:{PORTA}"
    print(f"\nAbrindo navegador em {url} …")
    webbrowser.open(url)

threading.Thread(target=_abrir_browser, daemon=True).start()

# ── Inicia o servidor ─────────────────────────────────────────────────────
os.chdir(PASTA)

print(f"\nServidor iniciando em http://localhost:{PORTA}")
print("Para acessar pelo celular/tablet na mesma rede Wi-Fi,")
print("use o endereço IP do seu computador, ex.: http://192.168.1.x:8000")
print("\nPressione Ctrl+C para encerrar o app.\n")

try:
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=PORTA, log_level="warning")
except KeyboardInterrupt:
    print("\nApp encerrado.")
except Exception as e:
    print(f"\n[ERRO] {e}")
    input("\nPressione Enter para fechar...")
