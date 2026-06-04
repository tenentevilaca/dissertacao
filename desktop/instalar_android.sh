#!/data/data/com.termux/files/usr/bin/bash
# ════════════════════════════════════════════════════════════════════
# INSTALAÇÃO — Verificador de Citações (Android / Termux)
# ════════════════════════════════════════════════════════════════════
# 1. Instale o Termux pelo F-Droid: https://f-droid.org/packages/com.termux/
# 2. Cole este arquivo no Termux e execute:  bash instalar_android.sh
# ════════════════════════════════════════════════════════════════════

set -e

echo ""
echo "=== Verificador de Citações — Instalação Android ==="
echo ""

echo "[1/5] Atualizando pacotes Termux..."
pkg update -y && pkg upgrade -y

echo ""
echo "[2/5] Instalando Python e dependências do sistema..."
pkg install -y python clang libzmq

echo ""
echo "[3/5] Instalando bibliotecas Python..."
pip install --upgrade pip
pip install flask anthropic pymupdf

echo ""
echo "[4/5] Concedendo acesso ao armazenamento externo (pen drive)..."
termux-setup-storage
echo "      → Quando o Android perguntar, toque em PERMITIR"
sleep 3

echo ""
echo "[5/5] Verificando instalação..."
python -c "import flask, anthropic; print('  ✓ Flask OK'); print('  ✓ Anthropic OK')"
python -c "import fitz; print('  ✓ PyMuPDF OK')" 2>/dev/null || echo "  ⚠ PyMuPDF não instalado (PDFs precisarão ser enviados nativamente à API)"

echo ""
echo "=== Instalação concluída! ==="
echo ""
echo "Para iniciar o verificador:"
echo "  1. Copie analisar.py e app_web.py para ~/:"
echo "     cp /storage/XXXX-XXXX/analisar.py ~/"
echo "     cp /storage/XXXX-XXXX/app_web.py ~/"
echo "  2. Execute: python ~/app_web.py"
echo "  3. Abra o Chrome e acesse: http://localhost:5000"
echo ""
