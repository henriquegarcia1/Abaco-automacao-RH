@echo off
rem Confere se o programa empacotado funciona neste computador (sem abrir a
rem tela). O resultado tambem fica salvo em verificacao.txt.
title Verificar Automacao Abaco
cd /d "%~dp0"
if not exist "%~dp0AutomacaoAbaco.exe" (
    echo [ERRO] AutomacaoAbaco.exe nao esta nesta pasta.
    pause
    exit /b 1
)
echo Verificando... aguarde alguns segundos.
start "" /wait "%~dp0AutomacaoAbaco.exe" --verificar
echo.
if exist "%~dp0verificacao.txt" (
    type "%~dp0verificacao.txt"
) else (
    echo Nao foi possivel gravar o resultado: a pasta pode estar protegida contra
    echo gravacao. Copie a pasta do programa para a Area de Trabalho ou Documentos.
)
echo.
pause
