@echo off
setlocal
rem Automacao Abaco + Excel: instala o que faltar (inclusive o Python) e abre o programa.
rem Basta dar dois cliques neste arquivo. Opcoes: -SoInstalar  -Reinstalar
title Automacao Abaco + Excel
pushd "%~dp0"

if not exist "%~dp0instalar.ps1" goto :sem_instalador

rem PowerShell do Windows pelo caminho completo; "Sysnative" garante o de 64 bits
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"

"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" %*
set "CODIGO=%ERRORLEVEL%"
popd

if "%CODIGO%"=="0" exit /b 0
echo.
echo A instalacao nao terminou. Leia a mensagem acima.
echo Se ela falar que a execucao de scripts esta desabilitada, peca ao suporte de TI
echo para liberar o PowerShell neste computador.
echo.
pause
exit /b %CODIGO%

:sem_instalador
echo.
echo [ERRO] O arquivo instalar.ps1 nao esta ao lado deste INICIAR.bat.
echo Se voce abriu direto de dentro de um arquivo .zip, clique com o botao direito
echo no .zip, escolha "Extrair tudo" e abra o INICIAR.bat da pasta extraida.
echo.
pause
exit /b 1
