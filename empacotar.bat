@echo off
setlocal
rem ==========================================================================
rem  EMPACOTAR - gera o programa que roda SEM Python instalado
rem ==========================================================================
rem  Resultado: dist\AutomacaoAbaco\AutomacaoAbaco.exe (mais a pasta _internal)
rem             dist\AutomacaoAbaco.zip  (o mesmo, compactado, para distribuir)
rem  Rode neste computador (que ja tem o ambiente do programa). Leva de 2 a 5
rem  minutos. Rode de novo sempre que o programa for atualizado.
rem ==========================================================================
title Empacotar Automacao Abaco
pushd "%~dp0"

echo [1/5] Conferindo o ambiente do programa (Python e bibliotecas)...
call "%~dp0INICIAR.bat" -SoInstalar
if errorlevel 1 goto :erro

set "PY=%LOCALAPPDATA%\AutomacaoAbaco\.venv\Scripts\python.exe"
if not exist "%PY%" (
    echo [ERRO] Ambiente do programa nao encontrado em %PY%
    goto :erro
)

echo.
echo [2/5] Instalando o PyInstaller no ambiente do programa...
"%PY%" -m pip install --disable-pip-version-check --prefer-binary --no-input "pyinstaller>=6.10"
if errorlevel 1 goto :erro

echo.
echo [3/5] Gerando o programa (2 a 5 minutos)...
if exist "%~dp0dist\AutomacaoAbaco" rmdir /s /q "%~dp0dist\AutomacaoAbaco"
"%PY%" -m PyInstaller --noconfirm --clean --distpath "%~dp0dist" --workpath "%TEMP%\AutomacaoAbaco_build" "%~dp0AutomacaoAbaco.spec"
if errorlevel 1 goto :erro

set "DEST=%~dp0dist\AutomacaoAbaco"
copy /y "%~dp0README.md" "%DEST%\" >nul
copy /y "%~dp0REGRAS_DE_NEGOCIO.md" "%DEST%\" >nul
copy /y "%~dp0VERIFICAR.bat" "%DEST%\" >nul
rem as configuracoes (endereco da Abaco, valores, regras) vao junto
if exist "%~dp0config.json" copy /y "%~dp0config.json" "%DEST%\" >nul

echo.
echo [4/5] Conferindo o programa gerado...
start "" /wait "%DEST%\AutomacaoAbaco.exe" --verificar
set "CODIGO=%ERRORLEVEL%"
if exist "%DEST%\verificacao.txt" (
    type "%DEST%\verificacao.txt"
    del "%DEST%\verificacao.txt"
)
if not "%CODIGO%"=="0" (
    echo [ERRO] A verificacao do programa gerado encontrou problemas. Veja acima.
    goto :erro
)

echo.
echo [5/5] Compactando para distribuir...
if exist "%~dp0dist\AutomacaoAbaco.zip" del "%~dp0dist\AutomacaoAbaco.zip"
powershell -NoProfile -Command "Compress-Archive -Path '%DEST%' -DestinationPath '%~dp0dist\AutomacaoAbaco.zip'"
if errorlevel 1 echo Nao consegui gerar o .zip; a pasta dist\AutomacaoAbaco esta pronta mesmo assim.

echo.
echo ==========================================================================
echo  PRONTO. Programa em: dist\AutomacaoAbaco\AutomacaoAbaco.exe
echo  Para outro computador: copie a pasta dist\AutomacaoAbaco inteira (ou o
echo  dist\AutomacaoAbaco.zip, extraido) para a Area de Trabalho ou Documentos
echo  e abra o AutomacaoAbaco.exe. Nao precisa instalar Python.
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
