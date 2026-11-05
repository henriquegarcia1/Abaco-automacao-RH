# -*- mode: python ; coding: utf-8 -*-
# ==========================================================================
#  Receita do PyInstaller: gera dist\AutomacaoAbaco.exe, UM ARQUIVO SÓ, que
#  roda SEM Python instalado. O Python, o Qt6, o pandas, o motor do Playwright
#  e o config.json padrão vão DENTRO do .exe (ao abrir, ele se descompacta
#  numa pasta temporária e apaga ao fechar). Use o empacotar.bat.
# ==========================================================================
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Playwright: o pacote PyInstaller não tem receita pronta para ele. O "motor"
# (driver: node + cli.js) é arquivo de dados da biblioteca e precisa ir junto.
datas = collect_data_files("playwright", include_py_files=False)

# config.json padrão (endereço da Ábaco, valores, regras). Na 1ª vez que o
# programa abre, ele é copiado para o lado do .exe, onde pode ser editado.
# fontes IBM Plex (visual do programa; licença SIL OFL em fontes/)
datas.append((os.path.join(SPECPATH, "fontes"), "fontes"))

_config = os.path.join(SPECPATH, "config.json")
if os.path.isfile(_config):
    datas.append((_config, "."))

hiddenimports = (
    ["nucleo_abaco", "regras_beneficios", "tema", "PySide6.QtSvg",
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
