@echo off
title Ferramentas Academicas - Verificador e Analisador
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo.
echo  =====================================================
echo   VERIFICADOR DE CITACOES ACADEMICAS
echo   Aguarde, inicializando...
echo  =====================================================
echo.

:: Tenta encontrar o Python de diferentes formas
set PYTHON_CMD=

:: Opção 1: comando "python"
python --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=python
    goto :encontrou
)

:: Opção 2: comando "python3"
python3 --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=python3
    goto :encontrou
)

:: Opção 3: Python Launcher do Windows (py)
py --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=py
    goto :encontrou
)

:: Opção 4: caminho comum do instalador oficial
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :encontrou
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :encontrou
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    goto :encontrou
)

:: Python não encontrado
echo [ERRO] Python nao encontrado no seu computador.
echo.
echo Para instalar, acesse:
echo   https://www.python.org/downloads/
echo.
echo Durante a instalacao, marque a opcao:
echo   "Add Python to PATH"
echo.
pause
exit /b 1

:encontrou
echo Python encontrado: %PYTHON_CMD%
echo.
%PYTHON_CMD% iniciar.py
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Algo deu errado. Verifique as mensagens acima.
    pause
)
