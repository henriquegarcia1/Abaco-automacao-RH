@echo off
setlocal
rem ==========================================================================
rem  EMPACOTAR - gera o programa que roda SEM Python instalado
rem ==========================================================================
rem  Resultado: dist\AutomacaoAbaco.exe  (UM arquivo so: Python, bibliotecas,
rem             motor do navegador e config.json padrao vao dentro dele)
rem  Rode neste computador (que ja tem o ambiente do programa). Leva de 3 a 6
rem  minutos. Rode de novo sempre que o programa ou o config.json mudar.
rem ==========================================================================
title Empacotar Automacao Abaco
pushd "%~dp0"

echo [1/4] Conferindo o ambiente do programa (Python e bibliotecas)...
call "%~dp0INICIAR.bat" -SoInstalar
if errorlevel 1 goto :erro

set "PY=%LOCALAPPDATA%\AutomacaoAbaco\.venv\Scripts\python.exe"
if not exist "%PY%" (
    echo [ERRO] Ambiente do programa nao encontrado em %PY%
    goto :erro
)

echo.
echo [2/4] Instalando o PyInstaller no ambiente do programa...
"%PY%" -m pip install --disable-pip-version-check --prefer-binary --no-input "pyinstaller>=6.10"
if errorlevel 1 goto :erro

echo.
echo [3/4] Gerando o programa em arquivo unico (3 a 6 minutos)...
set "SAIDA=%TEMP%\AutomacaoAbaco_dist"
if exist "%SAIDA%" rmdir /s /q "%SAIDA%"
"%PY%" -m PyInstaller --noconfirm --clean --distpath "%SAIDA%" --workpath "%TEMP%\AutomacaoAbaco_build" "%~dp0AutomacaoAbaco.spec"
if errorlevel 1 goto :erro
if not exist "%SAIDA%\AutomacaoAbaco.exe" (
    echo [ERRO] O PyInstaller nao gerou o AutomacaoAbaco.exe.
    goto :erro
)

echo.
echo [4/4] Conferindo o programa gerado...
rem confere numa pasta vazia: assim o teste usa o config.json de DENTRO do .exe
start "" /wait "%SAIDA%\AutomacaoAbaco.exe" --verificar
set "CODIGO=%ERRORLEVEL%"
if exist "%SAIDA%\verificacao.txt" type "%SAIDA%\verificacao.txt"
if not "%CODIGO%"=="0" (
    echo [ERRO] A verificacao do programa gerado encontrou problemas. Veja acima.
    goto :erro
)

if not exist "%~dp0dist" mkdir "%~dp0dist"
copy /y "%SAIDA%\AutomacaoAbaco.exe" "%~dp0dist\AutomacaoAbaco.exe" >nul
if errorlevel 1 (
    echo [ERRO] Nao consegui gravar dist\AutomacaoAbaco.exe. Feche o programa
    echo        se ele estiver aberto e rode de novo.
    goto :erro
)
rmdir /s /q "%SAIDA%" 2>nul

echo.
echo ==========================================================================
echo  PRONTO. Programa em: dist\AutomacaoAbaco.exe  (um arquivo so)
echo  Para outro computador: copie SO o AutomacaoAbaco.exe para a Area de
echo  Trabalho ou Documentos e abra. Nao precisa instalar Python.
echo  Na 1a vez ele cria, ao lado dele, o config.json (copia do embutido) e as
echo  pastas entrada e saida.
if exist "%~dp0dist\AutomacaoAbaco\AutomacaoAbaco.exe" (
    echo.
    echo  A pasta antiga dist\AutomacaoAbaco ^(com _internal^) nao e mais usada.
    echo  Antes de apaga-la, copie dela o que for seu ^(fluxos gravados, saida^).
)
if exist "%~dp0dist\config.json" (
    echo.
    echo  Atencao: dist\config.json existe e VALE MAIS que o embutido. Apague-o
    echo  se quiser usar o config.json novo que foi para dentro do .exe.
)
echo ==========================================================================
popd
pause
exit /b 0

:erro
echo.
echo O empacotamento nao terminou. Leia a mensagem acima.
popd
pause
exit /b 1
