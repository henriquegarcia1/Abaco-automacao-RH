# -*- coding: utf-8 -*-
"""Visual do programa (o mesmo do arquivo do Figma "Automação - Fluxos RH").

Cores, fontes (IBM Plex Sans e IBM Plex Mono, embutidas em fontes/) e a
folha de estilo do Qt. Cada setor tem a sua cor de destaque:
Benefícios = azul, Consulta de Admissões = areia, engrenagem = azul.
"""
import sys
import tempfile
from pathlib import Path

PAPEL = "#F5F7F9"
SUPERFICIE = "#FFFFFF"
TINTA = "#1C2733"
TINTA2 = "#556472"
TINTA3 = "#8695A3"
LINHA = "#DDE3E9"
LINHA_CAMPO = "#CBD3DB"
OK = "#2E7D4F"
ATENCAO = "#8A5A12"
ERRO = "#B83B3B"

SETORES = {
    # cor, cor ao passar o mouse, cor escura (texto), cor clara (fundo), selo
    "beneficios": ("#2B74B5", "#3482C6", "#1F5A8F", "#E7F0F8", "#2B74B5"),
    "admissoes": ("#7A6143", "#8A7052", "#5E4A32", "#F3ECE3", "#C9B79E"),
    "ajustes": ("#2B74B5", "#3482C6", "#1F5A8F", "#E7F0F8", "#8695A3"),
    "dev": ("#2B74B5", "#3482C6", "#1F5A8F", "#E7F0F8", "#8695A3"),
}

FAMILIA = "IBM Plex Sans"
FAMILIA_MONO = "IBM Plex Mono"


def _pasta_recursos():
    """fontes/ fica ao lado do programa (.py) ou dentro do .exe."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def carregar_fontes():
    """Registra as fontes embutidas. Sem elas, o Qt usa a do Windows."""
    from PySide6.QtGui import QFontDatabase
    pasta = _pasta_recursos() / "fontes"
    carregadas = 0
    for arq in sorted(pasta.glob("*.ttf")):
        if QFontDatabase.addApplicationFont(str(arq)) >= 0:
            carregadas += 1
    return carregadas


_IMAGENS = None


def _imagens():
    """Setas e marca de seleção em SVG (o Qt desenha pelos arquivos)."""
    global _IMAGENS
    if _IMAGENS:
        return _IMAGENS
    pasta = Path(tempfile.gettempdir()) / "AutomacaoAbaco_tema"
    pasta.mkdir(parents=True, exist_ok=True)
    svgs = {
        "seta": f'<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" '
                f'viewBox="0 0 12 12"><path d="M3 4.5l3 3 3-3" fill="none" '
                f'stroke="{TINTA2}" stroke-width="1.6" stroke-linecap="round" '
                f'stroke-linejoin="round"/></svg>',
        "seta_cima": f'<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" '
                     f'viewBox="0 0 12 12"><path d="M3 7.5l3-3 3 3" fill="none" '
                     f'stroke="{TINTA2}" stroke-width="1.6" stroke-linecap="round" '
                     f'stroke-linejoin="round"/></svg>',
        "marca": '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" '
                 'viewBox="0 0 12 12"><path d="M2.2 6.2l2.5 2.5 5-5.4" fill="none" '
                 'stroke="#ffffff" stroke-width="1.9" stroke-linecap="round" '
                 'stroke-linejoin="round"/></svg>',
    }
    _IMAGENS = {}
    for nome, svg in svgs.items():
        arq = pasta / f"{nome}.svg"
        try:
            arq.write_text(svg, encoding="utf-8")
        except OSError:
            pass
        _IMAGENS[nome] = arq.as_posix()
    return _IMAGENS


def estilo(setor="beneficios"):
    """Folha de estilo completa com a cor do setor."""
    cor, cor_mouse, escura, clara, _ = SETORES.get(setor, SETORES["beneficios"])
    img = _imagens()
    return f"""
* {{ font-family: "{FAMILIA}"; font-size: 13px; color: {TINTA}; }}
QMainWindow, QStackedWidget, QDialog {{ background: {PAPEL}; }}
QScrollArea {{ background: transparent; border: 0; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QToolTip {{ background: {TINTA}; color: #ffffff; border: 0; padding: 6px 8px; }}

/* textos */
QLabel#titulo {{ font-size: 24px; font-weight: 600; }}
QLabel#tituloCartao {{ font-size: 15px; font-weight: 600; }}
QLabel#numero {{ font-size: 22px; font-weight: 600; }}
QLabel#mesGrande {{ font-size: 34px; font-weight: 600; }}
QLabel#rotuloPequeno {{ font-size: 12px; color: {TINTA3}; }}
QLabel#periodoGrande {{ font-size: 26px; font-weight: 600; }}

/* blocos */
QFrame#cartao {{ background: {SUPERFICIE}; border: 1px solid {LINHA}; border-radius: 12px; }}
QFrame#cartao QLabel {{ background: transparent; }}
QFrame#destaque {{ background: {clara}; border: 1px solid {LINHA}; border-radius: 12px; }}
QFrame#destaque QLabel {{ background: transparent; }}
QFrame#linhaPasso {{ background: {LINHA}; border: 0; }}
QFrame#divisorFino {{ background: {LINHA}; border: 0; max-height: 1px; min-height: 1px; }}

/* botões */
QPushButton {{ background: {SUPERFICIE}; border: 1px solid {LINHA}; border-radius: 8px;
              padding: 7px 16px; font-weight: 600; }}
QPushButton:hover {{ background: #F3F6F9; border-color: {LINHA_CAMPO}; }}
QPushButton:pressed {{ background: #E9EEF2; }}
QPushButton:disabled {{ color: #A3AFBA; background: #F7F9FA; border-color: {LINHA}; }}
QPushButton#primario {{ background: {cor}; border: 1px solid {cor}; color: #ffffff; }}
QPushButton#primario:hover {{ background: {cor_mouse}; border-color: {cor_mouse}; }}
QPushButton#primario:pressed {{ background: {escura}; border-color: {escura}; }}
QPushButton#primario:disabled {{ background: #C9D2DA; border-color: #C9D2DA; color: #ffffff; }}
QPushButton#discreto {{ background: transparent; border: 1px solid transparent; color: {TINTA2};
    padding: 7px 10px; }}
QPushButton#discreto:hover {{ background: #EEF2F5; color: {TINTA}; }}
QPushButton#voltar {{ background: transparent; border: 0; color: {TINTA2}; padding: 6px 10px; }}
QPushButton#voltar:hover {{ background: #EEF2F5; color: {TINTA}; }}
QPushButton#grande {{ padding: 10px 18px; font-size: 14px; }}

/* campos */
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit, QPlainTextEdit, QTextEdit {{
    background: {SUPERFICIE}; border: 1px solid {LINHA_CAMPO}; border-radius: 8px;
    padding: 6px 10px; selection-background-color: {clara}; selection-color: {TINTA}; }}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QPlainTextEdit:focus, QTextEdit:focus {{ border-color: {cor}; }}
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled {{ color: #A3AFBA; background: #F7F9FA; }}
QComboBox {{ padding-right: 28px; }}
QComboBox::drop-down {{ border: 0; width: 26px; }}
QComboBox::down-arrow {{ image: url("{img['seta']}"); width: 12px; height: 12px; }}
QComboBox QAbstractItemView {{ background: {SUPERFICIE}; border: 1px solid {LINHA};
    selection-background-color: {clara}; selection-color: {TINTA}; outline: 0; padding: 4px; }}
QSpinBox, QDoubleSpinBox {{ padding-right: 22px; }}
QSpinBox::up-button, QDoubleSpinBox::up-button {{ subcontrol-position: top right; width: 20px;
    border: 0; border-top-right-radius: 8px; }}
QSpinBox::down-button, QDoubleSpinBox::down-button {{ subcontrol-position: bottom right; width: 20px;
    border: 0; border-bottom-right-radius: 8px; }}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{ image: url("{img['seta_cima']}"); width: 10px; height: 10px; }}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{ image: url("{img['seta']}"); width: 10px; height: 10px; }}
QPlainTextEdit#listaMatriculas {{ font-family: "{FAMILIA_MONO}"; font-size: 12px; background: #FDFCFA; }}

QCheckBox, QRadioButton {{ spacing: 9px; background: transparent; }}
QCheckBox::indicator {{ width: 16px; height: 16px; border-radius: 4px; border: 1px solid #9AA8B5;
    background: {SUPERFICIE}; }}
QCheckBox::indicator:hover {{ border-color: {cor}; }}
QCheckBox::indicator:checked {{ background: {cor}; border-color: {cor}; image: url("{img['marca']}"); }}
QRadioButton::indicator {{ width: 16px; height: 16px; border-radius: 8px; border: 1px solid #9AA8B5;
    background: {SUPERFICIE}; }}
QRadioButton::indicator:checked {{ border: 5px solid {cor}; }}

/* abas */
QTabWidget::pane {{ border: 0; border-top: 1px solid {LINHA}; top: -1px; background: transparent; }}
QTabBar {{ background: transparent; }}
QTabBar::tab {{ background: transparent; border: 0; border-bottom: 2px solid transparent;
    padding: 10px 14px; margin-right: 2px; color: {TINTA2}; }}
QTabBar::tab:hover {{ color: {TINTA}; }}
QTabBar::tab:selected {{ color: {escura}; border-bottom-color: {cor}; font-weight: 600; }}

/* tabelas e listas */
QTableView, QTableWidget, QTreeWidget, QTreeView, QListWidget {{ background: {SUPERFICIE};
    border: 1px solid {LINHA}; border-radius: 10px; gridline-color: #EEF1F4;
    alternate-background-color: #FAFBFC; selection-background-color: {clara};
    selection-color: {TINTA}; outline: 0; }}
QTableView::item, QTreeView::item {{ padding: 4px 6px; border: 0; }}
QHeaderView {{ background: transparent; }}
QHeaderView::section {{ background: #F9FAFB; color: {TINTA2}; font-weight: 600; font-size: 12px;
    border: 0; border-bottom: 1px solid {LINHA}; border-right: 1px solid #EEF1F4; padding: 8px 10px; }}
QTableCornerButton::section {{ background: #F9FAFB; border: 0; }}

/* casca do setor */
QFrame#topoSetor {{ background: {SUPERFICIE}; border: 0; border-bottom: 1px solid {LINHA}; }}
QFrame#selo {{ background: {SETORES.get(setor, SETORES['beneficios'])[4]}; border-radius: 3px; }}
QLabel#tituloSetor {{ font-size: 17px; font-weight: 600; background: transparent; }}
QFrame#separadorTopo {{ background: {LINHA}; }}
QListWidget#navegacao {{ background: {SUPERFICIE}; border: 0; border-right: 1px solid {LINHA};
    border-radius: 0; padding: 14px 10px; }}
QListWidget#navegacao::item {{ padding: 9px 12px; border-radius: 8px; color: {TINTA2}; margin-bottom: 2px; }}
QListWidget#navegacao::item:hover:!selected {{ background: #F3F6F9; color: {TINTA}; }}
QListWidget#navegacao::item:selected {{ background: {clara}; color: {escura}; font-weight: 600; }}
QListWidget#navegacao::item:disabled {{ background: transparent; color: {TINTA3};
    padding: 14px 12px 4px 12px; }}
QLabel#seloDev {{ background: #EEF1F4; color: {TINTA2}; border-radius: 11px; padding: 3px 10px;
    font-size: 12px; font-weight: 600; }}

/* registro e status */
QFrame#caixaRegistro {{ background: #FBFCFD; border: 0; border-top: 1px solid {LINHA}; }}
QLabel#tituloRegistro {{ font-weight: 600; background: transparent; }}
QPlainTextEdit#registro {{ font-family: "{FAMILIA_MONO}"; font-size: 12px; color: #3D4B58;
    background: #FBFCFD; border: 0; border-radius: 0; padding: 4px 8px; }}
QFrame#pergunta {{ background: {clara}; border: 2px solid {cor}; border-radius: 10px; }}
QFrame#pergunta QLabel {{ background: transparent; }}
QStatusBar {{ background: {SUPERFICIE}; border-top: 1px solid {LINHA}; }}
QStatusBar QLabel {{ color: {TINTA3}; font-size: 12px; }}
QStatusBar::item {{ border: 0; }}
QProgressBar {{ border: 0; background: #E7EDF2; border-radius: 3px; max-height: 6px; }}
QProgressBar::chunk {{ background: {cor}; border-radius: 3px; }}
QSplitter::handle {{ background: transparent; }}

QScrollBar:vertical {{ background: transparent; width: 11px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {LINHA_CAMPO}; border-radius: 3px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: #AEB9C4; }}
QScrollBar:horizontal {{ background: transparent; height: 11px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {LINHA_CAMPO}; border-radius: 3px; min-width: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

QMessageBox {{ background: {SUPERFICIE}; }}
QMenu {{ background: {SUPERFICIE}; border: 1px solid {LINHA}; padding: 4px; }}
QMenu::item {{ padding: 6px 18px; border-radius: 6px; }}
QMenu::item:selected {{ background: {clara}; }}

/* tela inicial */
QWidget#telaInicial {{ background: {PAPEL}; }}
QLabel#saudacao {{ font-size: 15px; color: {TINTA2}; background: transparent; }}
QLabel#contaLogin {{ color: {TINTA2}; font-size: 13px; }}
QLabel#dicaInicial {{ font-size: 12px; color: {TINTA3}; background: transparent; }}
QPushButton#cartaoAzul {{ background: #2B74B5; border: 0; border-radius: 18px; }}
QPushButton#cartaoAzul:hover {{ background: #3180C3; }}
QPushButton#cartaoAzul:pressed {{ background: #23649C; }}
QPushButton#cartaoAzul QLabel {{ color: #ffffff; background: transparent; }}
QLabel#cartaoAzulNome {{ font-size: 26px; font-weight: 500; }}
QPushButton#cartaoBege {{ background: #EDE6DD; border: 1px solid #D8CFC4; border-radius: 18px; }}
QPushButton#cartaoBege:hover {{ background: #F2ECE4; }}
QPushButton#cartaoBege:pressed {{ background: #E3DBD0; }}
QPushButton#cartaoBege QLabel {{ background: transparent; }}
QLabel#cartaoBegeNome {{ font-size: 26px; font-weight: 500; color: {TINTA}; }}
QLabel#cartaoAzulDesc {{ font-size: 13px; color: #DCE8F4; }}
QLabel#cartaoBegeDesc {{ font-size: 13px; color: #5E5140; }}
QToolButton#engrenagem {{ background: {SUPERFICIE}; border: 1px solid #E3E8EE; border-radius: 32px; }}
QToolButton#engrenagem:hover {{ background: #F5F8FB; }}
"""


def estilo_numero_passo(estado, setor="beneficios"):
    """Círculo numerado da trilha de passos: feito, agora ou a fazer."""
    cor = SETORES.get(setor, SETORES["beneficios"])[0]
    if estado == "feito":
        return (f"background:{OK}; border:2px solid {OK}; border-radius:16px; "
                f"color:#ffffff; font-weight:600; font-size:14px;")
    if estado == "agora":
        return (f"background:{cor}; border:2px solid {cor}; border-radius:16px; "
                f"color:#ffffff; font-weight:600; font-size:14px;")
    return (f"background:{SUPERFICIE}; border:2px solid {LINHA}; border-radius:16px; "
            f"color:{TINTA2}; font-weight:600; font-size:14px;")
