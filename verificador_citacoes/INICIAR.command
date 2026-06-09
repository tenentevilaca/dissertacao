#!/bin/bash
# Mac: clique duas vezes neste arquivo para iniciar o app

cd "$(dirname "$0")"

echo ""
echo "  ====================================================="
echo "   FERRAMENTAS ACADÊMICAS — Verificador + Analisador"
echo "   Aguarde, inicializando..."
echo "  ====================================================="
echo ""

# Encontra Python
PYTHON_CMD=""
for cmd in python3 python python3.13 python3.12 python3.11; do
    if command -v "$cmd" &>/dev/null; then
        VER=$("$cmd" -c "import sys; print(sys.version_info >= (3,11))" 2>/dev/null)
        if [ "$VER" = "True" ]; then
            PYTHON_CMD="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON_CMD" ]; then
    echo "[ERRO] Python 3.11+ não encontrado."
    echo ""
    echo "Instale em: https://www.python.org/downloads/"
    echo ""
    read -p "Pressione Enter para fechar..."
    exit 1
fi

echo "Python encontrado: $PYTHON_CMD"
echo ""
$PYTHON_CMD iniciar.py
