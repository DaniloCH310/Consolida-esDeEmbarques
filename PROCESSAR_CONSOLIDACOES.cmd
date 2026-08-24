@echo off
chcp 65001 > nul
title Consolidações - Faturamento de Transportes

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0AUTOMACAO\Consolidar.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo Processamento encerrado.
) else (
    echo O processamento terminou com uma ou mais pendências.
)
pause
exit /b %EXIT_CODE%
