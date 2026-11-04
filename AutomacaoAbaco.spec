# -*- mode: python ; coding: utf-8 -*-
# ==========================================================================
#  Receita do PyInstaller: gera dist\AutomacaoAbaco\AutomacaoAbaco.exe, que
#  roda SEM Python instalado (o Python, o Qt6, o pandas e o motor do
#  Playwright vão dentro da pasta). Use o empacotar.bat, não rode à mão.
# ==========================================================================
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Playwright: o pacote PyInstaller não tem receita pronta para ele. O "motor"
# (driver: node + cli.js) é arquivo de dados da biblioteca e precisa ir junto.
datas = collect_data_files("playwright", include_py_files=False)

hiddenimports = (
    ["nucleo_abaco", "regras_beneficios",
     # o pandas abre .xlsx/.xls/HTML com estas bibliotecas, sem import direto
     "openpyxl", "xlrd", "lxml", "lxml.etree", "lxml.html"]
    + collect_submodules("playwright")
)

a = Analysis(
    ["abaco_automacao.py"],
    pathex=["."],
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

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="AutomacaoAbaco",
    console=False,          # sem janela preta
    upx=False,              # UPX costuma disparar antivírus
    debug=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name="AutomacaoAbaco",
    upx=False,
)
