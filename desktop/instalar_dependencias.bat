@echo off
echo ================================================
echo  INSTALANDO DEPENDENCIAS
echo ================================================
echo.

python --version 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERRO: Python nao encontrado!
    echo Baixe em: https://www.python.org/downloads/
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo Instalando bibliotecas necessarias...
pip install PyMuPDF
pip install python-docx

echo.
echo Instalando biblioteca de IA Claude (opcional - para verificacao semantica):
pip install anthropic

echo.
echo ================================================
echo  INSTALACAO CONCLUIDA!
echo  Execute: iniciar.bat
echo ================================================
pause
