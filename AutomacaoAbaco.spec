# -*- mode: python ; coding: utf-8 -*-
# ==========================================================================
#  Receita do PyInstaller: gera dist\AutomacaoAbaco.exe, UM ARQUIVO SÓ, que
#  roda SEM Python instalado. O Python, o Qt6, o pandas, o motor do Playwright,
#  o config.json e os fluxos gravados vão DENTRO do .exe (ao abrir, ele se
#  descompacta numa pasta temporária e apaga ao fechar). Use o empacotar.bat.
# ==========================================================================
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Playwright: o pacote PyInstaller não tem receita pronta para ele. O "motor"
# (driver: node + cli.js) é arquivo de dados da biblioteca e precisa ir junto.
datas = collect_data_files("playwright", include_py_files=False)

# fontes IBM Plex (visual do programa; licença SIL OFL em fontes/)
datas.append((os.path.join(SPECPATH, "fontes"), "fontes"))

# ---- configuração e fluxos gravados vão DENTRO do .exe ------------------
# O config.json embutido é o da pasta MAIS os ajustes feitos na tela
# Configurações deste computador (pasta de dados do programa). Os fluxos
# gravados (só a navegação: o login é sempre feito à mão) também vão junto. Assim o .exe roda em
# outro computador sem nenhum arquivo ao lado dele.
import json, tempfile


def _mesclar(base, extra):
    for k, v in extra.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _mesclar(base[k], v)
        else:
            base[k] = v
    return base


import sys
sys.path.insert(0, SPECPATH)
from nucleo_abaco import sem_credenciais as _sem_credenciais

_dados = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
                      "AutomacaoAbaco", "dados")
_tmp = tempfile.mkdtemp(prefix="automacao_embutir_")
_cfg = {}
_config = os.path.join(SPECPATH, "config.json")
if os.path.isfile(_config):
    with open(_config, encoding="utf-8") as f:
        _cfg = json.load(f)
_ajustes = os.path.join(_dados, "ajustes.json")
if os.path.isfile(_ajustes):
    with open(_ajustes, encoding="utf-8") as f:
        _mesclar(_cfg, json.load(f))
    print("[embutir] ajustes de Configurações incluídos:", _ajustes)
# login nos sistemas é sempre à mão: nada de usuário/senha no .exe
for _k in ("siga_usuario", "siga_senha"):
    _cfg.pop(_k, None)
with open(os.path.join(_tmp, "config.json"), "w", encoding="utf-8") as f:
    json.dump(_cfg, f, indent=2, ensure_ascii=False)
datas.append((os.path.join(_tmp, "config.json"), "."))

_nomes = {_cfg.get("arquivo_fluxo") or "fluxo_gravado.py",
          _cfg.get("arquivo_fluxo_siga") or "fluxo_siga.py"}
for _n in sorted(_nomes):
    _base, _ext = os.path.splitext(_n)
    for _nome in (_n, _base + "_corrigido" + _ext):
        for _pasta in (_dados, SPECPATH):
            _arq = os.path.join(_pasta, _nome)
            if os.path.isfile(_arq):
                # vai uma cópia SEM o que foi digitado no login
                _pasta_fluxos = os.path.join(_tmp, "fluxos")
                os.makedirs(_pasta_fluxos, exist_ok=True)
                with open(_arq, encoding="utf-8", errors="replace") as f:
                    _texto = _sem_credenciais(f.read())
                _copia = os.path.join(_pasta_fluxos, _nome)
                with open(_copia, "w", encoding="utf-8") as f:
                    f.write(_texto)
                datas.append((_copia, "fluxos"))
                print("[embutir] fluxo (sem usuário e senha):", _arq)
                break

hiddenimports = (
    ["nucleo_abaco", "regras_beneficios", "tema", "PySide6.QtSvg",
     # login pelo KeyAuth (tela e assinatura das respostas)
     "autenticacao", "tela_login",
     "cryptography.hazmat.primitives.asymmetric.ed25519",
     # o pandas abre .xlsx/.xls/HTML com estas bibliotecas, sem import direto
     "openpyxl", "xlrd", "lxml", "lxml.etree", "lxml.html"]
    + collect_submodules("playwright")
)

a = Analysis(
    ["abaco_automacao.py"],
    pathex=[SPECPATH],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    # o que não é usado: deixa o pacote menor
    excludes=["tkinter", "matplotlib", "IPython", "pytest", "notebook",
              "PyInstaller", "setuptools", "pip"],
    noarchive=False,
)
pyz = PYZ(a.pure)

# arquivo único: binários e dados vão dentro do .exe (sem COLLECT/_internal)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="AutomacaoAbaco",
    console=False,          # sem janela preta
    upx=False,              # UPX costuma disparar antivírus
    debug=False,
    runtime_tmpdir=None,    # descompacta em %TEMP%
)
