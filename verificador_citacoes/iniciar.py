"""
Lançador — funciona tanto como script Python quanto como .exe (PyInstaller).
"""

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

PORTA = 8000

# Garante que as mensagens apareçam imediatamente no terminal (sem buffer)
try:
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)
except Exception:
    pass

# ── Detecta se está rodando como executável empacotado ──────────────────
FROZEN = getattr(sys, "frozen", False)

if FROZEN:
    # Recursos ficam em sys._MEIPASS; o .exe fica em sys.executable
    BUNDLE_DIR = Path(sys._MEIPASS)
    EXE_DIR    = Path(sys.executable).parent
else:
    BUNDLE_DIR = Path(__file__).parent
    EXE_DIR    = Path(__file__).parent

# Garante que os módulos do bundle sejam encontrados
os.chdir(BUNDLE_DIR)
sys.path.insert(0, str(BUNDLE_DIR))

# ── Cabeçalho ────────────────────────────────────────────────────────────
print("""
╔══════════════════════════════════════════════════════════╗
║   FERRAMENTAS ACADÊMICAS                                ║
║   Verificador de Citações + Analisador de Dissertação   ║
╚══════════════════════════════════════════════════════════╝
""")

# ── Instala dependências (apenas quando rodando como script) ─────────────
if not FROZEN:
    import subprocess
    req = BUNDLE_DIR / "requirements.txt"
    print("Verificando dependências (pode demorar na primeira vez)...")
    resultado = subprocess.run(
        [sys.executable, "-m", "pip", "install", "-q", "-r", str(req)],
        capture_output=True, text=True
    )
    if resultado.returncode != 0:
        print("[AVISO] Problema ao instalar dependências:")
        print(resultado.stderr[:400])
    else:
        print("Dependências prontas. ✓")

# ── Abre o navegador após 2 s ─────────────────────────────────────────────
def _abrir_browser():
    time.sleep(2)
    url = f"http://localhost:{PORTA}"
    print(f"Abrindo navegador em {url} …")
    webbrowser.open(url)

threading.Thread(target=_abrir_browser, daemon=True).start()

# ── Inicia o servidor ─────────────────────────────────────────────────────
print(f"Servidor iniciando em http://localhost:{PORTA}")
print("Para acessar pelo tablet/celular na mesma rede Wi-Fi,")
print(f"use: http://<IP-DO-COMPUTADOR>:{PORTA}")
print("\nPressione Ctrl+C ou feche esta janela para encerrar.\n")

try:
    import uvicorn
    if FROZEN:
        import app as app_module
        uvicorn.run(app_module.app, host="0.0.0.0", port=PORTA, log_level="warning")
    else:
        uvicorn.run("app:app", host="0.0.0.0", port=PORTA, log_level="warning")
except KeyboardInterrupt:
    print("\nApp encerrado.")
except Exception as exc:
    print(f"\n[ERRO] {exc}")
    if FROZEN:
        input("\nPressione Enter para fechar...")
