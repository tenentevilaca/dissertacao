@echo off
python analisar.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Erro ao executar. Verifique se o Python esta instalado.
    echo Execute primeiro: instalar_dependencias.bat
    pause
)
