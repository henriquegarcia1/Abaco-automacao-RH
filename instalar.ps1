# ==========================================================================
#  INSTALADOR E INICIALIZADOR — Automação Ábaco + Excel
# ==========================================================================
#  Chamado pelo INICIAR.bat. Pode rodar quantas vezes quiser: o que já
#  está instalado é pulado e o programa abre em poucos segundos.
#
#   1. Confere o ambiente do programa. Ele fica em
#      %LOCALAPPDATA%\AutomacaoAbaco\.venv (fora da pasta do programa), então
#      copiar a pasta para outro computador nunca leva um ambiente quebrado.
#      Se o ambiente faltar ou não funcionar, é recriado.
#   2. Para criar o ambiente, procura um Python 3.10 a 3.14 de 64 bits
#      (lançador "py", registro do Windows, pastas padrão e PATH). Se não
#      houver, instala o Python 3.12 só para este usuário (sem administrador):
#      primeiro pelo winget; se não der, pelo instalador oficial do
#      python.org, que só é executado se a assinatura digital for da
#      Python Software Foundation.
#   3. Instala as bibliotecas do requirements.txt e confere se carregam.
#      Baixa o navegador controlado (Chromium); se não conseguir, o
#      programa usa o Microsoft Edge (ou Chrome) do próprio Windows.
#   4. Abre a tela do programa e avisa se ela fechar com erro.
#
#  Opções:  -SoInstalar  instala e não abre
#           -Reinstalar  apaga o ambiente e instala tudo de novo
# ==========================================================================
param([switch]$SoInstalar, [switch]$Reinstalar)

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"   # deixa o download muito mais rápido
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
try { $Host.UI.RawUI.WindowTitle = "Automação Ábaco + Excel" } catch { }

# variáveis de outros programas que quebram qualquer Python
Remove-Item Env:PYTHONHOME, Env:PYTHONPATH -ErrorAction SilentlyContinue
$env:PYTHONIOENCODING = "utf-8"   # saída do pip legível nesta janela

$Pasta      = $PSScriptRoot
$Programa   = Join-Path $Pasta "abaco_automacao.py"
$Requisitos = Join-Path $Pasta "requirements.txt"
$VenvAntigo = Join-Path $Pasta ".venv"

$Local = $env:LOCALAPPDATA
if (-not $Local) { $Local = Join-Path $env:USERPROFILE "AppData\Local" }
$Dados   = Join-Path $Local "AutomacaoAbaco"
$Venv    = Join-Path $Dados ".venv"
$PyVenv  = Join-Path $Venv "Scripts\python.exe"
$PywVenv = Join-Path $Venv "Scripts\pythonw.exe"
$Marca   = Join-Path $Venv "instalado.txt"

$VersaoPython  = "3.12.10"
$UrlInstalador = "https://www.python.org/ftp/python/$VersaoPython/python-$VersaoPython-amd64.exe"
# versões aceitas, na ordem de preferência (todas têm as bibliotecas prontas)
$Preferencia = @("3.12", "3.13", "3.11", "3.14", "3.10")
$TesteBibliotecas = "import PySide6.QtWidgets, playwright.sync_api, pandas, openpyxl, xlrd, lxml.etree"

# --------------------------------------------------------------------------
function Passo($texto) { Write-Host ""; Write-Host ">> $texto" -ForegroundColor Cyan }
function Aviso($texto) { Write-Host "   $texto" -ForegroundColor Yellow }
function Falha($texto) {
    Write-Host ""
    Write-Host "[ERRO] $texto" -ForegroundColor Red
    exit 1
}

# Roda um Python e devolve versão, bits e caminho real; $null se não servir.
function Testar-Python([string]$Cmd, [string[]]$Extras) {
    if (-not $Extras) { $Extras = @() }
    $codigo = "import sys,struct;sys.stdout.reconfigure(encoding='utf-8');" +
              "print('PYOK|%d.%d|%d|%s' % (sys.version_info[0], sys.version_info[1], " +
              "struct.calcsize('P')*8, sys.executable))"
    try {
        $saida = & $Cmd @Extras -I -c $codigo 2>$null
    } catch {
        return $null
    }
    if ($LASTEXITCODE -ne 0) { return $null }
    $linha = @($saida) | Where-Object { "$_".StartsWith("PYOK|") } | Select-Object -Last 1
    if (-not $linha) { return $null }
    $partes = "$linha".Split("|")
    return [pscustomobject]@{
        Cmd = $Cmd; Extras = $Extras
        Versao = $partes[1]; Bits = [int]$partes[2]; Exe = $partes[3]
    }
}

function Candidatos-Python {
    $lista = New-Object System.Collections.ArrayList

    # 1. lançador "py" (no PATH ou nos lugares em que o instalador o coloca)
    $lancadores = @()
    $g = Get-Command py.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($g) { $lancadores += $g.Path }
    $lancadores += "$Local\Programs\Python\Launcher\py.exe"
    $lancadores += "$env:WINDIR\py.exe"
    foreach ($l in ($lancadores | Select-Object -Unique)) {
        if ($l -and (Test-Path -LiteralPath $l)) {
            foreach ($v in $Preferencia) {
                [void]$lista.Add([pscustomobject]@{ Cmd = $l; Extras = @("-$v") })
            }
        }
    }

    # 2. registro do Windows (instalações do python.org, para o usuário ou para todos)
    foreach ($raiz in "HKCU:\Software\Python\PythonCore", "HKLM:\Software\Python\PythonCore",
                      "HKLM:\Software\WOW6432Node\Python\PythonCore") {
        foreach ($chave in @(Get-ChildItem -Path $raiz -ErrorAction SilentlyContinue)) {
            $ip = Get-ItemProperty -LiteralPath "$($chave.PSPath)\InstallPath" -ErrorAction SilentlyContinue
            if (-not $ip) { continue }
            $exe = $ip.ExecutablePath
            if (-not $exe -and $ip.'(default)') { $exe = "$($ip.'(default)'.TrimEnd('\'))\python.exe" }
            if ($exe -and (Test-Path -LiteralPath $exe)) {
                [void]$lista.Add([pscustomobject]@{ Cmd = $exe; Extras = @() })
            }
        }
    }

    # 3. pastas padrão
    foreach ($padrao in "$Local\Programs\Python\Python3*\python.exe",
                        "$Local\Python\pythoncore-3*\python.exe",
                        "$env:ProgramFiles\Python3*\python.exe",
                        "$env:SystemDrive\Python3*\python.exe") {
        foreach ($f in @(Get-Item -Path $padrao -ErrorAction SilentlyContinue)) {
            [void]$lista.Add([pscustomobject]@{ Cmd = $f.FullName; Extras = @() })
        }
    }

    # 4. "python" no PATH (o atalho da Microsoft Store que não instala nada é descartado no teste)
    foreach ($g in @(Get-Command python.exe -All -ErrorAction SilentlyContinue)) {
        [void]$lista.Add([pscustomobject]@{ Cmd = $g.Path; Extras = @() })
    }
    return $lista
}

function Encontrar-Python {
    $achados = @{}
    foreach ($c in (Candidatos-Python)) {
        $r = Testar-Python $c.Cmd $c.Extras
        if (-not $r -or $r.Bits -ne 64 -or $Preferencia -notcontains $r.Versao) { continue }
        if (-not $achados.ContainsKey($r.Exe)) { $achados[$r.Exe] = $r }
    }
    foreach ($v in $Preferencia) {
        $m = @($achados.Values | Where-Object { $_.Versao -eq $v }) | Select-Object -First 1
        if ($m) { return $m }
    }
    return $null
}

function Atualizar-Path {
    $m = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $u = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$m;$u"
}

function Baixar([string]$Url, [string]$Destino) {
    try {   # proxy da rede, com o login do Windows
        $wp = [Net.WebRequest]::DefaultWebProxy
        if ($wp) { $wp.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials }
    } catch { }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destino -UseBasicParsing -ErrorAction Stop
        return $true
    } catch {
        Aviso "O download direto falhou ($($_.Exception.Message)). Tentando pelo BITS do Windows."
    }
    try {
        Start-BitsTransfer -Source $Url -Destination $Destino -ErrorAction Stop
        return $true
    } catch {
        Aviso "O BITS também falhou ($($_.Exception.Message))."
    }
    return $false
}

function Instalar-Python {
    if (-not [Environment]::Is64BitOperatingSystem) {
        Falha "Este programa precisa do Windows de 64 bits."
    }
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($winget) {
        Passo "Instalando o Python 3.12 pelo winget (pode levar alguns minutos)"
        & $winget.Path install --id Python.Python.3.12 --exact --scope user --silent --source winget `
            --accept-package-agreements --accept-source-agreements | Out-Host
        Atualizar-Path
        $py = Encontrar-Python
        if ($py) { return $py }
        Aviso "O winget não concluiu; tentando pelo site oficial."
    }

    Passo "Baixando o instalador oficial do Python $VersaoPython (python.org)"
    $destino = Join-Path $env:TEMP "python-$VersaoPython-amd64.exe"
    if (-not (Baixar $UrlInstalador $destino)) {
        Falha ("Não consegui baixar o Python. Verifique a internet ou instale à mão o Python 3.12 " +
               "de https://www.python.org/downloads/windows/ e abra o INICIAR.bat de novo.")
    }

    $assinatura = Get-AuthenticodeSignature -LiteralPath $destino
    if ($assinatura.Status -ne "Valid" -or
        $assinatura.SignerCertificate.Subject -notmatch "Python Software Foundation") {
        Remove-Item -LiteralPath $destino -Force -ErrorAction SilentlyContinue
        Falha "A assinatura digital do instalador não confere. Instalação interrompida por segurança."
    }

    Passo "Instalando o Python $VersaoPython só para este usuário (sem administrador)"
    $proc = Start-Process -FilePath $destino -Wait -PassThru -ArgumentList @(
        "/quiet", "InstallAllUsers=0", "InstallLauncherAllUsers=0", "PrependPath=1",
        "Include_launcher=1", "Include_pip=1", "Include_test=0", "Shortcuts=0")
    Remove-Item -LiteralPath $destino -Force -ErrorAction SilentlyContinue
    Atualizar-Path
    $py = Encontrar-Python
    if ($py) { return $py }
    if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
        Falha "O instalador do Python terminou com código $($proc.ExitCode)."
    }
    Falha "O Python foi instalado mas não foi encontrado. Feche esta janela e abra o INICIAR.bat de novo."
}

function Apagar-Pasta([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return $true }
    try { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop } catch { }
    if (Test-Path -LiteralPath $p) { cmd.exe /c rmdir /s /q $p 2>$null | Out-Null }
    return -not (Test-Path -LiteralPath $p)
}

# O ambiente existe, roda (o Python de base está neste computador) e tem pip?
function Ambiente-Funciona {
    if (-not (Test-Path -LiteralPath $PyVenv)) { return $false }
    $r = Testar-Python $PyVenv @()
    if (-not $r -or $r.Bits -ne 64 -or $Preferencia -notcontains $r.Versao) { return $false }
    & $PyVenv -I -m pip --version *> $null
    return ($LASTEXITCODE -eq 0)
}

function Bibliotecas-Carregam([switch]$MostrarErro) {
    if ($MostrarErro) { & $PyVenv -I -c $TesteBibliotecas | Out-Host }
    else { & $PyVenv -I -c $TesteBibliotecas *> $null }
    return ($LASTEXITCODE -eq 0)
}

function Rodar-Pip([string[]]$Argumentos) {
    $base = @("-m", "pip") + $Argumentos +
            @("--disable-pip-version-check", "--no-input", "--timeout", "60", "--retries", "5")
    & $PyVenv @base | Out-Host
    if ($LASTEXITCODE -eq 0) { return $true }
    Aviso "Tentando de novo com os certificados do Windows (redes com proxy)..."
    $base2 = $base + "--use-feature=truststore"
    & $PyVenv @base2 | Out-Host
    return ($LASTEXITCODE -eq 0)
}

function Navegador-Do-Windows {
    foreach ($c in "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                   "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                   "$Local\Microsoft\Edge\Application\msedge.exe",
                   "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                   "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                   "$Local\Google\Chrome\Application\chrome.exe") {
        if (Test-Path -LiteralPath $c) { return $true }
    }
    return $false
}

# ==========================================================================
Write-Host "==================================================================="
Write-Host "  AUTOMAÇÃO ÁBACO + EXCEL — verificando a instalação"
Write-Host "==================================================================="

if (-not (Test-Path -LiteralPath $Programa) -or -not (Test-Path -LiteralPath $Requisitos)) {
    Falha ("Não encontrei abaco_automacao.py e requirements.txt ao lado do INICIAR.bat. " +
           "Se abriu de dentro de um .zip, extraia a pasta inteira e abra o INICIAR.bat de lá.")
}

# ------------------------------------------------- 0. .venv antigo na pasta
# Versões anteriores criavam o ambiente dentro da pasta do programa; copiado
# para outro computador, ele aponta para um Python que não existe aqui.
if (Test-Path -LiteralPath (Join-Path $VenvAntigo "pyvenv.cfg")) {
    Passo "Removendo o .venv antigo da pasta do programa (o ambiente agora fica em $Venv)"
    if (-not (Apagar-Pasta $VenvAntigo)) {
        Aviso "Não consegui apagar a pasta .venv antiga. Ela não é mais usada; pode apagar à mão."
    }
}

# ------------------------------------------------- 1. ambiente + Python
if ($Reinstalar) {
    Passo "Apagando o ambiente atual para reinstalar"
    if (-not (Apagar-Pasta $Venv)) {
        Falha "Não consegui apagar $Venv. Feche o programa (se estiver aberto) e rode de novo."
    }
}

if (Ambiente-Funciona) {
    Write-Host "   Ambiente do programa OK."
} else {
    if (Test-Path -LiteralPath $Venv) {
        Aviso "O ambiente do programa está incompleto ou quebrado; vou recriar."
        if (-not (Apagar-Pasta $Venv)) {
            Falha "Não consegui apagar $Venv. Feche o programa (se estiver aberto) e rode de novo."
        }
    }

    Passo "Procurando o Python 3.10 a 3.14 (64 bits)"
    $py = Encontrar-Python
    if (-not $py) {
        Aviso "Nenhum Python compatível neste computador."
        $py = Instalar-Python
    }
    Write-Host "   Usando Python $($py.Versao): $($py.Exe)"

    Passo "Criando o ambiente isolado do programa"
    New-Item -ItemType Directory -Force -Path $Dados | Out-Null
    $extras = @($py.Extras)
    & $py.Cmd @extras -m venv $Venv | Out-Host
    if (-not (Ambiente-Funciona) -and (Test-Path -LiteralPath $PyVenv)) {
        & $PyVenv -I -m ensurepip --upgrade --default-pip | Out-Host   # venv criado sem pip
    }
    if (-not (Ambiente-Funciona)) { Falha "Não consegui criar o ambiente em $Venv." }
}

# ------------------------------------------------- 2. bibliotecas + navegador
$hash = (Get-FileHash -LiteralPath $Requisitos -Algorithm SHA256).Hash
$instalado = ""
if (Test-Path -LiteralPath $Marca) { $instalado = "$(Get-Content -LiteralPath $Marca -Raw)".Trim() }

if ($instalado -eq $hash -and (Bibliotecas-Carregam)) {
    Write-Host "   Bibliotecas já instaladas."
} else {
    Passo "Instalando as bibliotecas (Qt6, Playwright, pandas...). Na 1ª vez demora alguns minutos."
    Rodar-Pip @("install", "--upgrade", "pip") | Out-Null
    if (-not (Rodar-Pip @("install", "-r", $Requisitos, "--prefer-binary"))) {
        Falha "Não consegui instalar as bibliotecas. Verifique a internet e rode de novo."
    }
    if (-not (Bibliotecas-Carregam -MostrarErro)) {
        Falha ("As bibliotecas foram instaladas mas não carregam (erro acima). Se falar de DLL, " +
               "instale o 'Microsoft Visual C++ Redistributable x64' (https://aka.ms/vs/17/release/vc_redist.x64.exe) e rode de novo.")
    }

    Passo "Baixando o navegador controlado (Chromium)"
    & $PyVenv -m playwright install chromium | Out-Host
    if ($LASTEXITCODE -ne 0) {
        if (Navegador-Do-Windows) {
            Aviso "Não consegui baixar o Chromium. Sem problema: o programa vai usar o Edge/Chrome deste computador."
        } else {
            Falha "Não consegui baixar o navegador e não há Edge nem Chrome instalados. Verifique a internet e rode de novo."
        }
    }
    Set-Content -LiteralPath $Marca -Value $hash -Encoding ASCII
}

# ------------------------------------------------- 3. abrir
if ($SoInstalar) {
    Passo "Instalação concluída."
    exit 0
}

Passo "Abrindo o programa (esta janela fecha sozinha)"
if (-not (Test-Path -LiteralPath $PywVenv)) { $PywVenv = $PyVenv }
$proc = Start-Process -FilePath $PywVenv -ArgumentList "`"$Programa`"" -WorkingDirectory $Pasta -PassThru
$null = $proc.Handle   # guarda o processo para conseguir ler o código de saída
if ($proc.WaitForExit(6000) -and $proc.ExitCode -ne 0) {
    Aviso "O programa fechou logo ao abrir. Detalhes:"
    & $PyVenv $Programa
    Falha "O programa não abriu (veja a mensagem acima)."
}
exit 0
