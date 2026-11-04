#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================================
  AUTOMAÇÃO ÁBACO + EXCEL  —  tela do programa (Qt6)
==========================================================================

Como abrir: dois cliques no INICIAR.bat. Ele instala o que faltar
(inclusive o Python) e abre esta tela. Com tudo instalado, também dá:

    python abaco_automacao.py

Este arquivo só monta as telas. O que o programa faz está em:
  - nucleo_abaco.py       navegador, observação, coleta, robô de consulta
  - regras_beneficios.py  cálculo de VR/VA (regras em REGRAS_DE_NEGOCIO.md)

As tarefas que usam o navegador rodam em segundo plano. Quando o robô
precisa da pessoa ("faça login", "baixou?"), aparece um painel com botões
acima do registro; a tela nunca trava.
==========================================================================
"""

import copy
import math
import numbers
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

try:
    from PySide6.QtCore import (QAbstractTableModel, QModelIndex, QObject,
                                QSortFilterProxyModel, Qt, Signal)
    from PySide6.QtGui import QColor, QFont, QPalette
    from PySide6.QtWidgets import (
        QAbstractItemView, QApplication, QCheckBox, QComboBox, QFileDialog,
        QFormLayout, QFrame, QGridLayout, QHBoxLayout, QHeaderView,
        QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem,
        QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
        QScrollArea, QSpinBox, QSplitter, QStackedWidget, QStyle, QTableView,
        QTableWidget, QTableWidgetItem, QTabWidget, QTreeWidget,
        QTreeWidgetItem, QVBoxLayout, QWidget)
except ImportError:
    print("A tela precisa do Qt6 (PySide6), que ainda não está instalado.\n"
          "Dê dois cliques no INICIAR.bat: ele instala tudo e abre o programa.")
    sys.exit(1)

import nucleo_abaco as nucleo

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
         "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

ESTILO = """
QFrame#cartao { border: 1px solid palette(mid); border-radius: 8px; }
QFrame#pergunta { border: 2px solid palette(highlight); border-radius: 8px; }
QLabel#titulo { font-size: 20px; font-weight: 600; }
QLabel#tituloCartao { font-size: 14px; font-weight: 600; }
QLabel#numero { font-size: 22px; font-weight: 600; }
QPushButton { padding: 5px 12px; }
QPushButton#primario { background: palette(highlight); color: palette(highlighted-text);
                       border: 0; border-radius: 4px; padding: 7px 16px; font-weight: 600; }
QPushButton#primario:disabled { background: palette(mid); }
QListWidget#navegacao { font-size: 13px; }
QListWidget#navegacao::item { padding: 8px 6px; }
"""

# campos que aceitam nomes de coluna extras (Configurações > Colunas)
CAMPOS_COLUNAS = ["matricula", "nome", "cpf", "cargo", "cargo_comissionado",
                  "jornada", "secretaria", "vinculo", "data_admissao",
                  "data_exoneracao", "situacao", "tipo_afastamento",
                  "inicio_afastamento", "termino_afastamento", "salario", "tipo",
                  "periodo", "competencia", "data", "quantidade", "inicio", "fim",
                  "desconto", "reembolso", "observacao"]

# campos aceitos nos arquivos de importação (validados ao salvar)
CAMPOS_LAYOUT = {
    "vr": ["matricula", "nome", "tipo", "quantidade", "valor_dia",
           "valor_total", "secretaria", "cargo"],
    "va_credito": ["matricula", "nome", "valor"],
    "va_desconto": ["matricula", "nome", "desconto"],
}


# ==========================================================================
# Ponte entre a tarefa em segundo plano e a tela
# ==========================================================================
class PonteUI(QObject):
    """O `ui` que o núcleo recebe. Os sinais atravessam as threads com
    segurança; perguntar() espera a resposta dada na tela."""
    sig_log = Signal(str)
    sig_pergunta = Signal(str, list)
    sig_arquivo = Signal(str)
    sig_fim = Signal(object)

    def __init__(self):
        super().__init__()
        self.cancelado = False
        self._evento = threading.Event()
        self._resposta = None

    # ---- chamados pela tarefa (thread de fundo)
    def log(self, texto):
        self.sig_log.emit(str(texto))

    def _esperar(self):
        self._evento.wait()
        if self.cancelado or self._resposta is None:
            raise nucleo.Cancelado()
        return self._resposta

    def perguntar(self, texto, opcoes):
        if self.cancelado:
            raise nucleo.Cancelado()
        self._evento.clear()
        self.sig_pergunta.emit(texto, list(opcoes))
        return self._esperar()

    def escolher_arquivo(self, texto):
        if self.cancelado:
            raise nucleo.Cancelado()
        self._evento.clear()
        self.sig_arquivo.emit(texto)
        return self._esperar()

    # ---- chamados pela tela
    def responder(self, resposta):
        self._resposta = resposta
        self._evento.set()

    def cancelar(self):
        self.cancelado = True
        self._resposta = None
        self._evento.set()


class Sinal(QObject):
    """Leva um resultado de uma thread qualquer para a tela."""
    pronto = Signal(object)


def em_segundo_plano(funcao, ao_terminar):
    """Roda funcao() fora da tela e entrega o resultado em ao_terminar."""
    s = Sinal()
    s.pronto.connect(ao_terminar)
    em_segundo_plano.vivos.add(s)          # evita o coletor de lixo
    s.pronto.connect(lambda _: em_segundo_plano.vivos.discard(s))

    def alvo():
        try:
            r = funcao()
        except Exception:
            r = None
        s.pronto.emit(r)
    threading.Thread(target=alvo, daemon=True).start()


em_segundo_plano.vivos = set()


# ==========================================================================
# Peças reutilizáveis
# ==========================================================================
def secundario(label):
    pal = label.palette()
    pal.setColor(QPalette.ColorRole.WindowText,
                 pal.color(QPalette.ColorRole.PlaceholderText))
    label.setPalette(pal)
    return label


def texto(t, wrap=True, sec=False):
    lb = QLabel(t)
    lb.setWordWrap(wrap)
    lb.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    return secundario(lb) if sec else lb


def botao(t, ao_clicar=None, primario=False):
    b = QPushButton(t)
    if primario:
        b.setObjectName("primario")
    if ao_clicar:
        b.clicked.connect(ao_clicar)
    return b


def cartao(titulo, descricao=None):
    """Quadro com título; devolve (quadro, layout interno)."""
    q = QFrame()
    q.setObjectName("cartao")
    lay = QVBoxLayout(q)
    lay.setContentsMargins(14, 12, 14, 12)
    t = QLabel(titulo)
    t.setObjectName("tituloCartao")
    lay.addWidget(t)
    if descricao:
        lay.addWidget(texto(descricao, sec=True))
    return q, lay


def linha(*widgets, esticar=True):
    w = QWidget()
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 0, 0, 0)
    for x in widgets:
        h.addWidget(x)
    if esticar:
        h.addStretch(1)
    return w


def rolavel(conteudo):
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.Shape.NoFrame)
    area.setWidget(conteudo)
    return area


def formatar(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    s = str(v)
    return "" if s in ("nan", "NaN", "None", "NaT") else s


def moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


class ModeloTabela(QAbstractTableModel):
    """Mostra um DataFrame; linhas PENDENTE ficam destacadas."""
    COR_PENDENTE = QColor(255, 165, 0, 70)

    def __init__(self, df=None):
        super().__init__()
        self.df = df
        self._status = None

    def definir(self, df):
        self.beginResetModel()
        self.df = df
        cols = list(df.columns) if df is not None else []
        self._status = cols.index("Status") if "Status" in cols else None
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() or self.df is None else len(self.df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() or self.df is None else len(self.df.columns)

    def data(self, idx, role=Qt.ItemDataRole.DisplayRole):
        if not idx.isValid():
            return None
        v = self.df.iat[idx.row(), idx.column()]
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return formatar(v)
        if role == Qt.ItemDataRole.BackgroundRole and self._status is not None \
                and self.df.iat[idx.row(), self._status] == "PENDENTE":
            return self.COR_PENDENTE
        if role == Qt.ItemDataRole.TextAlignmentRole and isinstance(v, numbers.Number):
            return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return None

    def headerData(self, s, orient, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole or self.df is None:
            return None
        if orient == Qt.Orientation.Horizontal:
            return str(self.df.columns[s])
        return str(s + 1)


class TabelaDados(QWidget):
    """Tabela só de leitura, com filtro por texto e ordenação."""

    def __init__(self):
        super().__init__()
        self.modelo = ModeloTabela()
        self.proxy = QSortFilterProxyModel()
        self.proxy.setSourceModel(self.modelo)
        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)
        self.view = QTableView()
        self.view.setModel(self.proxy)
        self.view.setSortingEnabled(True)
        self.view.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
        self.view.setAlternatingRowColors(True)
        self.view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.view.horizontalHeader().setStretchLastSection(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.view)

    def definir(self, df):
        if df is not None:
            df = df[[c for c in df.columns if not str(c).startswith("_")]]
        self.modelo.definir(df)
        self.view.resizeColumnsToContents()
        for c in range(self.modelo.columnCount()):
            if self.view.columnWidth(c) > 320:
                self.view.setColumnWidth(c, 320)

    def filtrar(self, t):
        self.proxy.setFilterFixedString(t)


class TabelaEditavel(QWidget):
    """Tabela editável com botões de adicionar/remover linha."""

    def __init__(self, colunas, ajuda=None):
        super().__init__()
        self.t = QTableWidget(0, len(colunas))
        self.t.setHorizontalHeaderLabels(colunas)
        self.t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.t.verticalHeader().setVisible(False)
        self.t.setMinimumHeight(140)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        if ajuda:
            lay.addWidget(texto(ajuda, sec=True))
        lay.addWidget(self.t)
        lay.addWidget(linha(botao("+ Adicionar linha", self.adicionar),
                            botao("Remover selecionada", self.remover)))

    def adicionar(self, valores=None):
        r = self.t.rowCount()
        self.t.insertRow(r)
        for c in range(self.t.columnCount()):
            v = (valores or [""] * self.t.columnCount())[c]
            self.t.setItem(r, c, QTableWidgetItem("" if v is None else str(v)))
        if not valores:
            self.t.setCurrentCell(r, 0)
            self.t.editItem(self.t.item(r, 0))

    def remover(self):
        r = self.t.currentRow()
        if r >= 0:
            self.t.removeRow(r)

    def definir(self, linhas):
        self.t.setRowCount(0)
        for l in linhas:
            self.adicionar(l)

    def linhas(self):
        out = []
        for r in range(self.t.rowCount()):
            vals = [(self.t.item(r, c).text().strip() if self.t.item(r, c) else "")
                    for c in range(self.t.columnCount())]
            if any(vals):
                out.append(vals)
        return out


class PainelPergunta(QFrame):
    """Aparece quando a tarefa precisa de uma ação da pessoa."""

    def __init__(self):
        super().__init__()
        self.setObjectName("pergunta")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 10)
        topo = QLabel("A tarefa está esperando por você")
        topo.setObjectName("tituloCartao")
        self.msg = texto("")
        self.botoes = QHBoxLayout()
        lay.addWidget(topo)
        lay.addWidget(self.msg)
        lay.addLayout(self.botoes)
        self.hide()

    def mostrar(self, mensagem, opcoes, responder):
        while self.botoes.count():
            w = self.botoes.takeAt(0).widget()
            if w:
                w.deleteLater()
        self.msg.setText(mensagem)
        for i, op in enumerate(opcoes):
            b = botao(op, primario=(i == 0))
            b.clicked.connect(lambda _=False, o=op: (self.hide(), responder(o)))
            self.botoes.addWidget(b)
        self.botoes.addStretch(1)
        self.show()


# ==========================================================================
# Telas
# ==========================================================================
class Pagina(QWidget):
    titulo = ""
    descricao = ""
    expandir = False    # True: o último bloco ocupa a altura que sobrar

    def __init__(self, janela):
        super().__init__()
        self.j = janela
        self.corpo = QVBoxLayout()
        self.corpo.setSpacing(12)
        conteudo = QWidget()
        lay = QVBoxLayout(conteudo)
        lay.setContentsMargins(20, 16, 20, 16)
        t = QLabel(self.titulo)
        t.setObjectName("titulo")
        lay.addWidget(t)
        if self.descricao:
            lay.addWidget(texto(self.descricao, sec=True))
        lay.addLayout(self.corpo, 1)
        if not self.expandir:
            lay.addStretch(1)
        externo = QVBoxLayout(self)
        externo.setContentsMargins(0, 0, 0, 0)
        externo.addWidget(rolavel(conteudo))

    def ao_mostrar(self):
        pass


# --------------------------------------------------------------------------
class PaginaInicio(Pagina):
    titulo = "Automação Ábaco + Excel"
    descricao = ("O robô prepara, a pessoa confere. Nada vai para a Ábaco sem a "
                 "conferência de quem é responsável pelo benefício.")

    def __init__(self, janela):
        super().__init__(janela)
        q, lay = cartao("Situação")
        self.grade = QGridLayout()
        lay.addLayout(self.grade)
        lay.addWidget(linha(botao("Atualizar", self.ao_mostrar)))
        self.corpo.addWidget(q)

        q, lay = cartao("Rotina do mês — benefícios")
        self.rotina = texto("")
        lay.addWidget(self.rotina)
        lay.addWidget(linha(
            botao("Ir para Benefícios VR/VA", lambda: self.j.ir("beneficios"), True),
            botao("Configurações", lambda: self.j.ir("config"))))
        self.corpo.addWidget(q)

        q, lay = cartao("Diagnóstico de processo",
                        "Antes de automatizar uma tarefa nova, observe e grave como "
                        "a pessoa trabalha.")
        lay.addWidget(linha(botao("Ir para Diagnóstico", lambda: self.j.ir("diagnostico"))))
        self.corpo.addWidget(q)

    def _item(self, r, nome, ok, detalhe):
        simbolo = "✔" if ok is True else ("…" if ok is None else "⚠")
        rot = QLabel(f"{simbolo}  {nome}")
        fonte = rot.font()
        fonte.setBold(True)
        rot.setFont(fonte)
        self.grade.addWidget(rot, r, 0)
        det = texto(detalhe, sec=True)
        self.grade.addWidget(det, r, 1)
        return rot, det

    def ao_mostrar(self):
        while self.grade.count():
            w = self.grade.takeAt(0).widget()
            if w:
                w.deleteLater()
        cfg = self.j.cfg
        libs = nucleo.bibliotecas_instaladas()
        faltam = [k for k, v in libs.items() if not v]
        r = 0
        erro_cfg = nucleo.erro_no_config()
        if erro_cfg:
            self._item(r, "config.json", False, f"Com erro, usando o padrão: {erro_cfg}")
            r += 1
        self._item(r, "Bibliotecas", not faltam,
                   "Todas instaladas." if not faltam else
                   f"Faltam: {', '.join(faltam)}. Abra a tela Instalação.")
        r += 1
        nav_rot, nav = self._item(r, "Navegador controlado", None, "verificando...")
        r += 1
        padrao = cfg["url_abaco"] == nucleo.PADRAO["url_abaco"]
        self._item(r, "Endereço da Ábaco", not padrao,
                   "Ainda é o endereço de exemplo. Ajuste em Configurações > Geral."
                   if padrao else cfg["url_abaco"])
        r += 1
        if not libs.get("pandas"):
            self._item(r, "Valores de VR/VA", False, "Instale as bibliotecas primeiro.")
        else:
            regras = nucleo.regras_do_config(cfg)
            vals = {t: v for t, v in regras["vr"]["valor_dia_por_tipo"].items()
                    if v not in (None, "")}
            va = regras["va"].get("valor_credito")
            ok = bool(vals) and (va not in (None, "") or not regras["va"].get("ativo"))
            self._item(r, "Valores de VR/VA", ok,
                       "VR por dia: " + ", ".join(f"{t} {moeda(float(v))}"
                                                  for t, v in vals.items())
                       + (f"; VA {moeda(float(va))}" if va not in (None, "") else
                          "; VA sem valor") + ". Confira em Configurações > Valores."
                       if ok or vals else
                       "Ainda não informados: as planilhas saem só com quantidades. "
                       "Configurações > Valores.")
        r += 1
        sel = nucleo.seletores_configurados(cfg)
        self._item(r, "Robô de consulta", sel,
                   "Seletores configurados." if sel else
                   "Ainda não sabe onde ficam os campos na tela da Ábaco. "
                   "Grave uma consulta e preencha em Configurações > Robô de consulta.")
        self.grade.setColumnStretch(1, 1)

        def mostrar_nav(situacao):
            ok, texto = situacao or (False, "Não consegui verificar.")
            try:
                nav.setText(texto)
                nav_rot.setText(("✔" if ok else "⚠") + "  Navegador controlado")
            except RuntimeError:
                pass    # a tela foi redesenhada enquanto verificava
        if libs.get("playwright"):
            em_segundo_plano(nucleo.situacao_navegador, mostrar_nav)
        else:
            mostrar_nav((False, "Instale as bibliotecas primeiro (INICIAR.bat)."))

        prazos = (cfg.get("beneficios") or {}).get("prazos") or \
            {"documentacao_ate_dia": 5, "folha_fecha_dia": 20}
        self.rotina.setText(
            f"1. Até o dia {prazos['documentacao_ate_dia']}: chega a documentação; "
            f"a secretaria envia os plantões.\n"
            f"2. Coletar os relatórios da competência (SIGA e Ábaco).\n"
            f"3. Conferir os novos admitidos na Ábaco (robô).\n"
            f"4. Calcular VR/VA e resolver as pendências.\n"
            f"5. Importar na Ábaco antes do dia {prazos['folha_fecha_dia']} "
            f"(fechamento da folha).\n"
            f"6. Enviar a planilha de novos admitidos à Educação.")


# --------------------------------------------------------------------------
class PaginaBeneficios(Pagina):
    titulo = "Benefícios VR/VA"
    descricao = ("Escolha a competência (o mês do benefício; o pagamento é "
                 "antecipado) e siga os três passos.")

    def __init__(self, janela):
        super().__init__(janela)
        ano, mes = nucleo.competencia_padrao()
        self.mes = QComboBox()
        self.mes.addItems(MESES)
        self.mes.setCurrentIndex(mes - 1)
        self.ano = QSpinBox()
        self.ano.setRange(2020, 2100)
        self.ano.setValue(ano)
        self.mes.currentIndexChanged.connect(self.atualizar_dias)
        self.ano.valueChanged.connect(self.atualizar_dias)
        self.mes.currentIndexChanged.connect(self.atualizar_entrada)
        self.ano.valueChanged.connect(self.atualizar_entrada)
        self.corpo.addWidget(linha(QLabel("Competência:"), self.mes, self.ano))

        # ---- passo 1
        q, lay = cartao("1. Coletar relatórios",
                        "A base é a planilha de carga do mês anterior (a aba CARGA, "
                        "com o tipo de cada servidor); a partir do 2º mês o programa "
                        "usa a que ele mesmo gerou. Depois o navegador guia a emissão "
                        "do relatório do SIGA e dos de férias e faltas da Ábaco, EM "
                        "EXCEL. No fim, pede a planilha de plantões que a secretaria "
                        "manda por e-mail.")
        self.arquivos = QTreeWidget()
        self.arquivos.setHeaderLabels(["Arquivo", "Origem", "Situação"])
        self.arquivos.setRootIsDecorated(False)
        self.arquivos.setMinimumHeight(150)
        self.arquivos.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.arquivos.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.arquivos)
        lay.addWidget(linha(
            botao("Coletar relatórios", self.coletar, True),
            botao("Adicionar arquivo manualmente...", self.adicionar_arquivo),
            botao("Abrir pasta de entrada", self.abrir_entrada),
            botao("Gerar modelos vazios", self.gerar_modelos)))
        self.corpo.addWidget(q)

        # ---- passo 2
        q, lay = cartao("2. Conferir novos admitidos na Ábaco (opcional)",
                        "O robô digita cada matrícula nova na tela de consulta da "
                        "Ábaco e lê data de admissão, cargo, jornada e secretaria. "
                        "No cálculo, os dados da Ábaco prevalecem e cada diferença "
                        "vira pendência.")
        self.aviso_robo = texto("", sec=True)
        lay.addWidget(self.aviso_robo)
        self.bt_conferir = botao("Conferir na Ábaco", self.conferir, True)
        lay.addWidget(linha(self.bt_conferir,
                            botao("Configurar o robô", lambda: self.j.ir("config", "robo"))))
        self.corpo.addWidget(q)

        # ---- passo 3
        q, lay = cartao("3. Calcular VR/VA",
                        "Atualiza a planilha do mês anterior: VR = (dias úteis − "
                        "férias − faltas) × valor do tipo; plantonista = plantões × "
                        "valor. Aplica afastamentos, exonerações, limite salarial do "
                        "diferenciado e inclui os novos admitidos (com complementar). "
                        "VA: crédito e desconto; quem faltou perde. Gera a planilha no "
                        "formato da Divisão (ex.: '11- NOVEMBRO - 2026'), os arquivos "
                        "de importação e um modelo de ajustes para as pendências.")
        self.dias = QSpinBox()
        self.dias.setRange(0, 31)
        self.info_dias = texto("", sec=True)
        lay.addWidget(linha(QLabel("Dias úteis da competência:"), self.dias))
        lay.addWidget(self.info_dias)
        lay.addWidget(linha(botao("Calcular VR/VA", self.calcular, True)))
        self.corpo.addWidget(q)

    def competencia(self):
        return self.ano.value(), self.mes.currentIndex() + 1

    def ao_mostrar(self):
        self.atualizar_entrada()
        self.atualizar_dias()
        sel = nucleo.seletores_configurados(self.j.cfg)
        self.bt_conferir.setEnabled(sel)
        self.aviso_robo.setText(
            "Robô configurado." if sel else
            "⚠ O robô ainda não sabe onde ficam os campos na tela da Ábaco. Grave "
            "uma consulta em Diagnóstico > Gravar fluxo e preencha os seletores em "
            "Configurações > Robô de consulta.")

    def atualizar_entrada(self):
        self.arquivos.clear()
        ano, mes = self.competencia()
        for nome, sistema, descricao, arq, quando in nucleo.status_entrada(
                self.j.cfg, ano, mes):
            situacao = f"{arq.name} — {quando:%d/%m/%Y %H:%M}" if arq else "falta"
            if arq and nome == "carga_anterior" and arq.parent != \
                    nucleo.pasta_entrada(self.j.cfg):
                situacao = f"{arq.name} (gerada no mês anterior)"
            if not arq and nome in nucleo.OPCIONAIS:
                situacao = nucleo.OPCIONAIS[nome]
            it = QTreeWidgetItem([descricao, sistema, situacao])
            it.setToolTip(0, descricao)
            self.arquivos.addTopLevelItem(it)

    def atualizar_dias(self):
        ano, mes = self.competencia()
        try:
            du, folgas, sugestoes = nucleo.info_dias_uteis(self.j.cfg, ano, mes)
        except ImportError:
            self.info_dias.setText("Instale as bibliotecas para calcular os dias úteis.")
            return
        except Exception as e:
            self.info_dias.setText(f"⚠ {e}")
            return
        self.dias.setValue(du)
        partes = [f"Calculado: {du} dias úteis."]
        if folgas:
            partes.append("Não contam: " + "; ".join(f"{d:%d/%m} {desc}" for d, desc in folgas) + ".")
        if sugestoes:
            partes.append("⚠ Pontos facultativos comuns NÃO descontados: "
                          + "; ".join(f"{d:%d/%m} {desc}" for d, desc in sugestoes)
                          + ". Se o decreto do município liberar, cadastre em "
                            "Configurações > Calendário ou ajuste o número acima.")
        self.info_dias.setText(" ".join(partes))

    def coletar(self):
        ano, mes = self.competencia()
        cfg = self.j.cfg
        self.j.executar("Coletar relatórios",
                        lambda ui: nucleo.coletar_relatorios(cfg, ano, mes, ui),
                        lambda _: self.atualizar_entrada())

    def adicionar_arquivo(self):
        nomes = {"Planilha de carga do mês anterior": "carga_anterior",
                 "Servidores (SIGA - situação de afastamento)": "servidores",
                 "Férias (Ábaco)": "ferias", "Faltas (Ábaco)": "faltas",
                 "Plantões (secretaria)": "plantoes",
                 "Ajustes manuais (tipo, quantidade, desconto, reembolso)": "ajustes"}
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Escolha o relatório", str(Path.home() / "Downloads"),
            "Planilhas (*.xlsx *.xls *.csv *.xlsm);;Todos (*.*)")
        if not caminho:
            return
        tipo, ok = QInputDialog.getItem(self, "Qual relatório é este?",
                                        "Tipo de relatório:", list(nomes), 0, False)
        if ok:
            destino = nucleo.copiar_para_entrada(caminho, self.j.cfg, nomes[tipo])
            self.j.log(f"Arquivo copiado para {destino}")
            self.atualizar_entrada()

    def abrir_entrada(self):
        p = nucleo.pasta_entrada(self.j.cfg)
        p.mkdir(parents=True, exist_ok=True)
        nucleo.abrir_no_sistema(p)

    def gerar_modelos(self):
        try:
            criados = nucleo.gerar_modelos()
        except ImportError:
            QMessageBox.warning(self, "Faltam bibliotecas", "Abra a tela Instalação.")
            return
        self.j.log("Modelos criados: " + ", ".join(c.name for c in criados))
        nucleo.abrir_no_sistema(criados[0].parent)

    def conferir(self):
        ano, mes = self.competencia()
        cfg = self.j.cfg
        self.j.executar("Conferir na Ábaco",
                        lambda ui: nucleo.conferir_na_abaco(cfg, ano, mes, ui),
                        lambda _: self.atualizar_entrada())

    def calcular(self):
        ano, mes = self.competencia()
        cfg, du = self.j.cfg, self.dias.value()
        self.j.executar("Calcular VR/VA",
                        lambda ui: nucleo.calcular_beneficios(cfg, ano, mes, du, ui),
                        self.j.mostrar_resultado)


# --------------------------------------------------------------------------
class PaginaResultados(Pagina):
    titulo = "Resultados"
    descricao = ("Confira tudo antes de importar. Linhas em laranja estão "
                 "pendentes e NÃO foram para o arquivo de importação.")
    expandir = True

    ABAS = [("Pendências", "pendencias"), ("VR", "vr"), ("VA", "va"),
            ("Complementar", "complementar"), ("Excluídos", "excluidos"),
            ("Aguardando próximo ciclo", "aguardando"), ("Educação", "educacao"),
            ("Fora da planilha", "fora")]

    def __init__(self, janela):
        super().__init__(janela)
        self.res = None
        self.arquivos = []
        self.cabecalho = texto("Nenhum cálculo feito ainda. Vá em Benefícios VR/VA "
                               "> Calcular.")
        self.corpo.addWidget(self.cabecalho)

        self.numeros = QHBoxLayout()
        self.corpo.addLayout(self.numeros)
        self.avisos = texto("", sec=True)
        self.corpo.addWidget(self.avisos)

        self.busca = QLineEdit()
        self.busca.setPlaceholderText("Filtrar por matrícula, nome, motivo...")
        self.busca.textChanged.connect(self.filtrar)
        self.abas = QTabWidget()
        self.tabelas = {}
        for rotulo, chave in self.ABAS:
            t = TabelaDados()
            self.tabelas[chave] = t
            self.abas.addTab(t, rotulo)
        self.abas.setMinimumHeight(380)
        self.abas.currentChanged.connect(lambda _: self.filtrar(self.busca.text()))
        self.corpo.addWidget(self.busca)
        self.corpo.addWidget(self.abas, 1)

        self.bt_conf = botao("Abrir planilha de conferência", self.abrir_conferencia, True)
        self.bt_pasta = botao("Abrir pasta da competência", self.abrir_pasta)
        self.bt_conf.setEnabled(False)
        self.bt_pasta.setEnabled(False)
        self.corpo.addWidget(linha(self.bt_conf, self.bt_pasta,
                                   botao("Voltar aos passos", lambda: self.j.ir("beneficios"))))

    def _numero(self, valor, rotulo):
        q = QFrame()
        q.setObjectName("cartao")
        lay = QVBoxLayout(q)
        n = QLabel(str(valor))
        n.setObjectName("numero")
        lay.addWidget(n)
        lay.addWidget(texto(rotulo, sec=True))
        return q

    def mostrar(self, res, arquivos):
        self.res, self.arquivos = res, arquivos
        vr_ok = int((res.vr["Status"] == "OK").sum())
        va_ok = int((res.va["Crédito"].fillna(0) > 0).sum())
        trava = int((res.pendencias["Trava a importação"] == "Sim").sum())
        total_vr = res.vr.loc[res.vr["Status"] == "OK", "Valor total"].fillna(0).sum()
        af, mf = res.folha
        self.cabecalho.setText(
            f"Competência {res.mes:02d}/{res.ano} (folha {mf:02d}/{af}) · "
            f"{res.dias_uteis} dias úteis · VR prontos: {moeda(float(total_vr))} · "
            f"gerado às {datetime.now():%H:%M}")
        while self.numeros.count():
            w = self.numeros.takeAt(0).widget()
            if w:
                w.deleteLater()
        for valor, rotulo in [
                (f"{vr_ok}/{len(res.vr)}", "VR prontos para importar"),
                (f"{va_ok}", "créditos de VA"),
                (f"{len(res.pendencias)}", f"pendências ({trava} travam)"),
                (len(res.complementar), "na complementar"),
                (len(res.educacao), "novos da Educação")]:
            self.numeros.addWidget(self._numero(valor, rotulo))
        self.avisos.setText("\n".join(f"• {a}" for a in res.avisos))
        for _, chave in self.ABAS:
            self.tabelas[chave].definir(getattr(res, chave))
        self.abas.setCurrentIndex(0 if len(res.pendencias) else 1)
        self.bt_conf.setEnabled(True)
        self.bt_pasta.setEnabled(True)

    def filtrar(self, t):
        w = self.abas.currentWidget()
        if w:
            w.filtrar(t)

    def abrir_conferencia(self):
        if self.arquivos:
            nucleo.abrir_no_sistema(self.arquivos[0])

    def abrir_pasta(self):
        if self.arquivos:
            nucleo.abrir_no_sistema(self.arquivos[0].parent)


# --------------------------------------------------------------------------
class PaginaDiagnostico(Pagina):
    titulo = "Diagnóstico de processo"
    descricao = ("Antes de automatizar uma tarefa nova: observe como a pessoa "
                 "trabalha, decida o que vale automatizar e grave o fluxo.")

    def __init__(self, janela):
        super().__init__(janela)
        q, lay = cartao(
            "Observar o trabalho",
            "Abre a Ábaco. A pessoa faz login e trabalha normalmente; o programa "
            "registra as telas, quando se copia/cola e de qual campo, mas NUNCA o "
            "conteúdo copiado nem a senha. Há um aviso '🔴 Observando' na tela o "
            "tempo todo. Ao fechar o navegador, sai um relatório.")
        self.bt_rel = botao("Abrir último relatório", self.abrir_relatorio)
        md = nucleo.PASTA / f"{self.j.cfg['relatorio']}.md"
        self.bt_rel.setEnabled(md.exists())
        lay.addWidget(linha(botao("Começar a observar", self.observar, True), self.bt_rel))
        self.corpo.addWidget(q)

        q, lay = cartao(
            "Gravar fluxo",
            "A pessoa executa UMA tarefa do início ao fim e o gravador do Playwright "
            "escreve os passos em código. Use também para ensinar o robô de consulta: "
            "grave a pesquisa de uma matrícula na tela de consulta do servidor.")
        lay.addWidget(linha(botao("Gravar fluxo", self.gravar, True),
                            botao("Abrir fluxo gravado", self.abrir_fluxo)))
        self.corpo.addWidget(q)

        q, lay = cartao(
            "Extrair tabela para Excel",
            "Lê a tabela que está na tela da Ábaco (várias páginas seguidas) e grava "
            "numa planilha.")
        lay.addWidget(linha(botao("Extrair tabela", self.extrair, True)))
        self.corpo.addWidget(q)

    def observar(self):
        cfg = self.j.cfg

        def fim(md):
            self.bt_rel.setEnabled(True)
            if md and QMessageBox.question(self, "Relatório pronto",
                                           "Abrir o relatório da observação?") \
                    == QMessageBox.StandardButton.Yes:
                nucleo.abrir_no_sistema(md)
        self.j.executar("Observar", lambda ui: nucleo.observar(cfg, ui), fim)

    def gravar(self):
        cfg = self.j.cfg
        self.j.executar("Gravar fluxo", lambda ui: nucleo.gravar(cfg, ui))

    def extrair(self):
        cfg = self.j.cfg

        def fim(xlsx):
            if xlsx and QMessageBox.question(self, "Extração pronta",
                                             f"Abrir {xlsx.name}?") \
                    == QMessageBox.StandardButton.Yes:
                nucleo.abrir_no_sistema(xlsx)
        self.j.executar("Extrair tabela", lambda ui: nucleo.extrair(cfg, ui), fim)

    def abrir_relatorio(self):
        md = nucleo.PASTA / f"{self.j.cfg['relatorio']}.md"
        if md.exists():
            nucleo.abrir_no_sistema(md)

    def abrir_fluxo(self):
        f = nucleo.PASTA / self.j.cfg["arquivo_fluxo"]
        if f.exists():
            nucleo.abrir_no_sistema(f)
        else:
            QMessageBox.information(self, "Sem fluxo", "Nenhum fluxo gravado ainda.")


# --------------------------------------------------------------------------
class PaginaConfig(Pagina):
    titulo = "Configurações"
    descricao = ("Tudo fica salvo no config.json, na pasta do programa. Os itens "
                 "marcados 'a validar' em REGRAS_DE_NEGOCIO.md se ajustam aqui.")
    expandir = True

    def __init__(self, janela):
        super().__init__(janela)
        self.abas = QTabWidget()
        self.abas.setMinimumHeight(420)
        self.indices = {}
        self.corpo.addWidget(self.abas, 1)

        # ---- Geral
        w = QWidget()
        f = QFormLayout(w)
        self.url = QLineEdit()
        self.mostrar = QCheckBox("Mostrar o navegador enquanto o robô trabalha "
                                 "(recomendado)")
        self.entrada = QLineEdit()
        self.saida = QLineEdit()
        self.fluxo = QLineEdit()
        self.planilha = QLineEdit()
        self.relatorio = QLineEdit()
        f.addRow("Endereço (URL) da Ábaco:", self.url)
        f.addRow("", self.mostrar)
        f.addRow("Pasta de entrada (relatórios):", self.entrada)
        f.addRow("Pasta de saída (planilhas):", self.saida)
        f.addRow("Arquivo do fluxo gravado:", self.fluxo)
        f.addRow("Planilha da extração:", self.planilha)
        f.addRow("Nome do relatório de observação:", self.relatorio)
        self._aba(w, "Geral", "geral")

        # ---- Valores
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(QLabel("Vale-refeição: valor por DIA de cada tipo (R7)"))
        self.jornadas = TabelaEditavel(
            ["Tipo", "Valor por dia (R$)"],
            "Os tipos da aba CARGA: NORMAL (40h), PLANTÃO (12h), DIFERENCIADO (6h) "
            "e DIFERENCIADO 2 (4h30).")
        v.addWidget(self.jornadas)
        f1 = QFormLayout()
        self.limite = QLineEdit()
        self.tipos_limite = QLineEdit()
        self.plantao_padrao = QSpinBox()
        self.plantao_padrao.setRange(0, 31)
        f1.addRow("Limite salarial do diferenciado (R$) (R8):", self.limite)
        f1.addRow("Tipos com limite salarial:", self.tipos_limite)
        f1.addRow("Plantões de um plantonista novo (R6):", self.plantao_padrao)
        v.addLayout(f1)
        v.addWidget(QLabel("Vale-alimentação"))
        f2 = QFormLayout()
        self.va_ativo = QCheckBox("Calcular vale-alimentação (VA)")
        self.va_educa = QCheckBox("Incluir os servidores da Educação no VA (A5)")
        self.va_valor = QLineEdit()
        self.va_desc = QLineEdit()
        self.va_afast = QLineEdit()
        self.va_falta = QComboBox()
        for chave, rot in [("perde", "Perde o VA no mês, sem o desconto (POP)"),
                           ("revisar", "Marcar para revisão"),
                           ("manter", "Mantém o VA")]:
            self.va_falta.addItem(rot, chave)
        f2.addRow("", self.va_ativo)
        f2.addRow("", self.va_educa)
        f2.addRow("Crédito do VA (R$) (A1):", self.va_valor)
        f2.addRow("Desconto do VA (R$) (A3):", self.va_desc)
        f2.addRow("Servidor com falta (A2):", self.va_falta)
        f2.addRow("Afastamentos que tiram o VA (A6):", self.va_afast)
        v.addLayout(f2)
        self._aba(w, "Valores", "valores")

        # ---- Calendário
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(texto("Os feriados nacionais (inclusive Sexta-feira Santa e "
                          "Consciência Negra) já são automáticos. Aqui entram os do "
                          "município e os pontos facultativos dos decretos.", sec=True))
        v.addWidget(QLabel("Feriados municipais"))
        self.feriados = TabelaEditavel(["Data (DD/MM/AAAA)", "Descrição"])
        v.addWidget(self.feriados)
        v.addWidget(QLabel("Pontos facultativos"))
        self.pontos = TabelaEditavel(["Data (DD/MM/AAAA)", "Descrição"])
        v.addWidget(self.pontos)
        self._aba(w, "Calendário", "calendario")

        # ---- Regras
        w = QWidget()
        v = QVBoxLayout(w)
        f3 = QFormLayout()
        self.educa = QLineEdit()
        self.afast = QLineEdit()
        self.sit_excluir = QLineEdit()
        self.sit_verificar = QLineEdit()
        f3.addRow("Lotação da Educação contém (D1):", self.educa)
        f3.addRow("Afastamentos que tiram o VR (E4):", self.afast)
        f3.addRow("Situações que excluem:", self.sit_excluir)
        f3.addRow("Situações para verificar (E5):", self.sit_verificar)
        v.addLayout(f3)
        v.addWidget(texto("Separe vários trechos por vírgula. A comparação ignora "
                          "acentos e maiúsculas e vale para trechos ('afastad' pega "
                          "afastado e afastada).", sec=True))
        dica = ("Campo: vinculo (o REGIME do SIGA), cargo, secretaria, situacao... "
                "'Ano de admissão' restringe ao ano da nomeação (ex.: PSS de 2022).")
        v.addWidget(QLabel("Quem não recebe VR (E1)"))
        self.exclusoes = TabelaEditavel(
            ["Campo", "Contém", "Ano de admissão (opcional)", "Motivo"], dica)
        v.addWidget(self.exclusoes)
        v.addWidget(QLabel("Quem não recebe VA (E1–E3)"))
        self.exclusoes_va = TabelaEditavel(
            ["Campo", "Contém", "Ano de admissão (opcional)", "Motivo"], dica)
        v.addWidget(self.exclusoes_va)
        f4 = QFormLayout()
        self.doc_dia = QSpinBox()
        self.doc_dia.setRange(1, 31)
        self.folha_dia = QSpinBox()
        self.folha_dia.setRange(1, 31)
        f4.addRow("Documentação até o dia (C1):", self.doc_dia)
        f4.addRow("Folha fecha no dia (C2):", self.folha_dia)
        v.addLayout(f4)
        self._aba(w, "Regras", "regras")

        # ---- Importação
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(texto("Colunas dos arquivos importados na Ábaco, na ordem. "
                          "Confirme o modelo com a Ábaco (R9).", sec=True))
        v.addWidget(QLabel("Vale-refeição"))
        self.layout_vr = TabelaEditavel(
            ["Título da coluna", "Campo"],
            "Campos: " + ", ".join(CAMPOS_LAYOUT["vr"]))
        v.addWidget(self.layout_vr)
        v.addWidget(QLabel("Vale-alimentação: crédito"))
        self.layout_va = TabelaEditavel(
            ["Título da coluna", "Campo"],
            "Campos: " + ", ".join(CAMPOS_LAYOUT["va_credito"]))
        v.addWidget(self.layout_va)
        v.addWidget(QLabel("Vale-alimentação: desconto"))
        self.layout_va_desc = TabelaEditavel(
            ["Título da coluna", "Campo"],
            "Campos: " + ", ".join(CAMPOS_LAYOUT["va_desconto"]))
        v.addWidget(self.layout_va_desc)
        self._aba(w, "Importação", "importacao")

        # ---- Robô de consulta
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(texto(
            "Onde ficam os campos na tela de consulta do servidor na Ábaco. Para "
            "descobrir: Diagnóstico > Gravar fluxo, pesquise UMA matrícula e veja no "
            "fluxo gravado os seletores (ex.: #txtMatricula). Ou mande o arquivo para "
            "quem dá suporte ao programa.", sec=True))
        f5 = QFormLayout()
        self.sel_mat = QLineEdit()
        self.sel_btn = QLineEdit()
        self.sel_btn.setPlaceholderText("vazio = tecla Enter no campo da matrícula")
        self.espera = QSpinBox()
        self.espera.setRange(0, 30000)
        self.espera.setSingleStep(250)
        self.espera.setSuffix(" ms")
        f5.addRow("Campo da matrícula:", self.sel_mat)
        f5.addRow("Botão pesquisar:", self.sel_btn)
        f5.addRow("Espera após pesquisar:", self.espera)
        self.sel_campos = {}
        for chave, rot in [("nome", "Nome"), ("data_admissao", "Data de admissão"),
                           ("cargo", "Cargo"), ("jornada", "Jornada"),
                           ("secretaria", "Secretaria/lotação")]:
            self.sel_campos[chave] = QLineEdit()
            f5.addRow(f"{rot}:", self.sel_campos[chave])
        v.addLayout(f5)
        v.addStretch(1)
        self._aba(w, "Robô de consulta", "robo")

        # ---- Colunas
        w = QWidget()
        v = QVBoxLayout(w)
        self.colunas = TabelaEditavel(
            ["Campo", "Nomes aceitos (separados por vírgula)"],
            "Só se algum relatório usar um nome de coluna que o programa não "
            "reconhece. Campos: " + ", ".join(CAMPOS_COLUNAS) + ".")
        v.addWidget(self.colunas)
        self._aba(w, "Colunas", "colunas")

        self.corpo.addWidget(linha(
            botao("Salvar", self.salvar, True),
            botao("Descartar alterações", self.carregar),
            botao("Abrir config.json", lambda: nucleo.abrir_no_sistema(nucleo.CONFIG)
                  if nucleo.CONFIG.exists() else None)))
        self.carregar()

    def _aba(self, w, rotulo, chave):
        self.indices[chave] = self.abas.addTab(rolavel(w), rotulo)

    def abrir_aba(self, chave):
        if chave in self.indices:
            self.abas.setCurrentIndex(self.indices[chave])

    # ---- config -> tela
    def carregar(self):
        cfg = self.j.cfg
        self.url.setText(cfg["url_abaco"])
        self.mostrar.setChecked(bool(cfg["mostrar_navegador"]))
        self.entrada.setText(cfg["pasta_entrada"])
        self.saida.setText(cfg["pasta_saida"])
        self.fluxo.setText(cfg["arquivo_fluxo"])
        self.planilha.setText(cfg["planilha"])
        self.relatorio.setText(cfg["relatorio"])

        tela = cfg["tela_consulta"]
        self.sel_mat.setText(tela.get("campo_matricula", ""))
        self.sel_btn.setText(tela.get("botao_pesquisar", ""))
        self.espera.setValue(int(tela.get("espera_ms", 1500)))
        for k, campo in self.sel_campos.items():
            campo.setText(tela.get("campos", {}).get(k, ""))

        try:
            regras = nucleo.regras_do_config(cfg)
        except ImportError:
            for chave in ("valores", "calendario", "regras", "importacao", "colunas"):
                self.abas.setTabEnabled(self.indices[chave], False)
            return
        dec = lambda v: "" if v in (None, "") else formatar(float(v))
        self.jornadas.definir([[k, dec(v)] for k, v in
                               regras["vr"]["valor_dia_por_tipo"].items()])
        self.limite.setText(dec(regras["vr"].get("limite_salarial")))
        self.tipos_limite.setText(", ".join(regras["vr"].get("tipos_com_limite_salarial") or []))
        self.plantao_padrao.setValue(int(regras["vr"].get("plantao_padrao") or 0))
        self.va_ativo.setChecked(bool(regras["va"].get("ativo")))
        self.va_educa.setChecked(bool(regras["va"].get("inclui_educacao")))
        self.va_valor.setText(dec(regras["va"].get("valor_credito")))
        self.va_desc.setText(dec(regras["va"].get("valor_desconto")))
        self.va_afast.setText(", ".join(regras["va"].get("afastamentos_sem_direito") or []))
        i = self.va_falta.findData(regras["va"].get("com_falta", "perde"))
        self.va_falta.setCurrentIndex(max(i, 0))
        br = lambda iso: datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y") \
            if len(iso) == 10 and iso[4] == "-" else iso
        self.feriados.definir([[br(d), desc] for d, desc in
                               regras["feriados_municipais"].items()])
        self.pontos.definir([[br(d), desc] for d, desc in
                             regras["pontos_facultativos"].items()])
        self.educa.setText(regras["secretaria_educacao_contem"])
        self.afast.setText(", ".join(regras["afastamentos_sem_direito"]))
        self.sit_excluir.setText(", ".join(regras["situacoes_excluir"]))
        self.sit_verificar.setText(", ".join(regras["situacoes_verificar"]))
        linhas_excl = lambda lista: [[e.get("campo", ""), e.get("contem", ""),
                                      e.get("ano_admissao") or "", e.get("motivo", "")]
                                     for e in lista]
        self.exclusoes.definir(linhas_excl(regras["vr"]["exclusoes"]))
        self.exclusoes_va.definir(linhas_excl(regras["va"]["exclusoes"]))
        self.doc_dia.setValue(int(regras["prazos"]["documentacao_ate_dia"]))
        self.folha_dia.setValue(int(regras["prazos"]["folha_fecha_dia"]))
        self.layout_vr.definir(regras["vr"]["layout_importacao"])
        self.layout_va.definir(regras["va"]["layout_credito"])
        self.layout_va_desc.definir(regras["va"]["layout_desconto"])
        self.colunas.definir([[k, ", ".join(v)] for k, v in
                              regras["colunas_extras"].items()])

    # ---- tela -> config
    def salvar(self):
        try:
            cfg = self._montar()
        except ValueError as e:
            QMessageBox.warning(self, "Confira as configurações", str(e))
            return
        nucleo.salvar_config(cfg)
        self.j.cfg = nucleo.carregar_config()
        self.j.log("Configurações salvas.")
        self.j.statusBar().showMessage("Configurações salvas.", 4000)

    def _montar(self):
        cfg = copy.deepcopy(self.j.cfg)
        if not self.url.text().strip():
            raise ValueError("Informe o endereço da Ábaco.")
        cfg.update({
            "url_abaco": self.url.text().strip(),
            "mostrar_navegador": self.mostrar.isChecked(),
            "pasta_entrada": self.entrada.text().strip() or "entrada",
            "pasta_saida": self.saida.text().strip() or "saida",
            "arquivo_fluxo": self.fluxo.text().strip() or "fluxo_gravado.py",
            "planilha": self.planilha.text().strip() or "dados.xlsx",
            "relatorio": self.relatorio.text().strip() or "relatorio_observacao",
        })
        cfg["tela_consulta"] = {
            "campo_matricula": self.sel_mat.text().strip(),
            "botao_pesquisar": self.sel_btn.text().strip(),
            "espera_ms": self.espera.value(),
            "campos": {k: c.text().strip() for k, c in self.sel_campos.items()},
        }
        if not self.abas.isTabEnabled(self.indices["valores"]):
            return cfg          # sem pandas: as regras não foram carregadas

        def numero(txt, onde):
            txt = txt.strip().replace("R$", "").strip()
            if not txt:
                return None
            if "," in txt:
                txt = txt.replace(".", "").replace(",", ".")
            try:
                return float(txt)
            except ValueError:
                raise ValueError(f"{onde}: '{txt}' não é um valor válido.")

        def data_iso(txt, onde):
            for formato in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(txt, formato).strftime("%Y-%m-%d")
                except ValueError:
                    continue
            raise ValueError(f"{onde}: data '{txt}' inválida. Use DD/MM/AAAA.")

        def lista(txt):
            return [t.strip() for t in txt.split(",") if t.strip()]

        import regras_beneficios as rb
        b = rb.regras_completas(cfg.get("beneficios"))
        tipos = {t.strip().upper(): numero(v, f"Valor do tipo {t}")
                 for t, v in self.jornadas.linhas() if t.strip()}
        if not tipos:
            raise ValueError("Valores: informe ao menos um tipo de VR.")
        b["vr"]["valor_dia_por_tipo"] = tipos
        b["vr"]["limite_salarial"] = numero(self.limite.text(), "Limite salarial")
        com_limite = [t.upper() for t in lista(self.tipos_limite.text())]
        for t in com_limite:
            if t not in tipos:
                raise ValueError(f"Tipos com limite salarial: '{t}' não é um dos "
                                 f"tipos ({', '.join(tipos)}).")
        b["vr"]["tipos_com_limite_salarial"] = com_limite
        b["vr"]["plantao_padrao"] = self.plantao_padrao.value()
        b["va"]["ativo"] = self.va_ativo.isChecked()
        b["va"]["inclui_educacao"] = self.va_educa.isChecked()
        b["va"]["valor_credito"] = numero(self.va_valor.text(), "Crédito do VA")
        b["va"]["valor_desconto"] = numero(self.va_desc.text(), "Desconto do VA")
        b["va"]["com_falta"] = self.va_falta.currentData()
        b["va"]["afastamentos_sem_direito"] = lista(self.va_afast.text())
        b["feriados_municipais"] = {data_iso(d, "Feriados"): desc
                                    for d, desc in self.feriados.linhas() if d}
        b["pontos_facultativos"] = {data_iso(d, "Pontos facultativos"): desc
                                    for d, desc in self.pontos.linhas() if d}
        b["secretaria_educacao_contem"] = self.educa.text().strip()
        b["afastamentos_sem_direito"] = lista(self.afast.text())
        b["situacoes_excluir"] = lista(self.sit_excluir.text())
        b["situacoes_verificar"] = lista(self.sit_verificar.text())

        def exclusoes(tabela, onde):
            saida = []
            for campo, contem, ano, motivo in tabela.linhas():
                if campo not in rb.COLUNAS:
                    raise ValueError(f"{onde}: campo '{campo}' não existe. Use um "
                                     f"de: {', '.join(rb.COLUNAS)}.")
                if not contem:
                    raise ValueError(f"{onde}: preencha 'Contém' na linha do "
                                     f"campo {campo}.")
                e = {"campo": campo, "contem": contem,
                     "motivo": motivo or f"{campo} contém '{contem}'"}
                if ano:
                    if not ano.isdigit():
                        raise ValueError(f"{onde}: ano '{ano}' inválido.")
                    e["ano_admissao"] = int(ano)
                saida.append(e)
            return saida
        b["vr"]["exclusoes"] = exclusoes(self.exclusoes, "Quem não recebe VR")
        b["va"]["exclusoes"] = exclusoes(self.exclusoes_va, "Quem não recebe VA")
        b["prazos"] = {"documentacao_ate_dia": self.doc_dia.value(),
                       "folha_fecha_dia": self.folha_dia.value()}
        for chave, tabela, beneficio, destino in (
                ("vr", self.layout_vr, "vr", "layout_importacao"),
                ("va_credito", self.layout_va, "va", "layout_credito"),
                ("va_desconto", self.layout_va_desc, "va", "layout_desconto")):
            lay = tabela.linhas()
            rotulo = chave.upper().replace("_", " ")
            for titulo_col, campo in lay:
                if campo not in CAMPOS_LAYOUT[chave]:
                    raise ValueError(f"Importação {rotulo}: campo '{campo}' "
                                     f"inválido. Use um de: "
                                     f"{', '.join(CAMPOS_LAYOUT[chave])}.")
            if not lay:
                raise ValueError(f"Importação {rotulo}: informe ao menos uma coluna.")
            b[beneficio][destino] = [[t or c, c] for t, c in lay]
        extras = {}
        for campo, nomes in self.colunas.linhas():
            if campo not in rb.COLUNAS:
                raise ValueError(f"Colunas: campo '{campo}' não existe.")
            extras[campo] = lista(nomes)
        b["colunas_extras"] = extras
        cfg["beneficios"] = b
        return cfg


# --------------------------------------------------------------------------
class PaginaInstalacao(Pagina):
    titulo = "Instalação"
    descricao = ("O Python, esta tela e as bibliotecas são instalados pelo "
                 "INICIAR.bat (dois cliques). Aqui dá para conferir e reinstalar "
                 "as bibliotecas e o navegador.")

    def __init__(self, janela):
        super().__init__(janela)
        q, lay = cartao("Situação")
        self.lista = QTreeWidget()
        self.lista.setHeaderLabels(["Item", "Situação"])
        self.lista.setRootIsDecorated(False)
        self.lista.setMinimumHeight(220)
        self.lista.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.lista)
        lay.addWidget(linha(botao("Instalar / atualizar bibliotecas e navegador",
                                  self.instalar, True),
                            botao("Verificar de novo", self.ao_mostrar)))
        self.corpo.addWidget(q)

    def ao_mostrar(self):
        self.lista.clear()
        self.lista.addTopLevelItem(QTreeWidgetItem(
            ["Python", f"{sys.version.split()[0]} — {sys.executable}"]))
        for pacote, ok in nucleo.bibliotecas_instaladas().items():
            self.lista.addTopLevelItem(QTreeWidgetItem(
                [pacote, "✔ instalado" if ok else "⚠ falta"]))
        nav = QTreeWidgetItem(["Navegador controlado", "verificando..."])
        self.lista.addTopLevelItem(nav)

        def mostrar(situacao):
            ok, texto = situacao or (False, "Não consegui verificar.")
            try:
                nav.setText(1, ("✔ " if ok else "⚠ ") + texto)
            except RuntimeError:
                pass
        em_segundo_plano(nucleo.situacao_navegador, mostrar)

    def instalar(self):
        self.j.executar("Instalação", nucleo.instalar_dependencias,
                        lambda _: self.ao_mostrar())


# ==========================================================================
# Janela principal
# ==========================================================================
class Janela(QMainWindow):
    PAGINAS = [("inicio", "Início", QStyle.StandardPixmap.SP_DirHomeIcon, PaginaInicio),
               ("beneficios", "Benefícios VR/VA", QStyle.StandardPixmap.SP_FileDialogDetailedView,
                PaginaBeneficios),
               ("resultados", "Resultados", QStyle.StandardPixmap.SP_FileDialogContentsView,
                PaginaResultados),
               ("diagnostico", "Diagnóstico", QStyle.StandardPixmap.SP_FileDialogInfoView,
                PaginaDiagnostico),
               ("config", "Configurações", QStyle.StandardPixmap.SP_FileDialogListView,
                PaginaConfig),
               ("instalacao", "Instalação", QStyle.StandardPixmap.SP_DriveHDIcon,
                PaginaInstalacao)]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automação Ábaco + Excel")
        self.resize(1180, 820)
        self.cfg = nucleo.carregar_config()
        self.ponte = None
        self.tarefa = None

        self.nav = QListWidget()
        self.nav.setObjectName("navegacao")
        self.nav.setFixedWidth(210)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.pilha = QStackedWidget()
        self.paginas = {}
        for chave, rotulo, icone, classe in self.PAGINAS:
            self.nav.addItem(QListWidgetItem(self.style().standardIcon(icone), rotulo))
            pagina = classe(self)
            self.paginas[chave] = pagina
            self.pilha.addWidget(pagina)
        self.nav.currentRowChanged.connect(self._trocar)

        # registro + painel de pergunta
        self.pergunta = PainelPergunta()
        self.registro = QPlainTextEdit()
        self.registro.setReadOnly(True)
        self.registro.setMaximumBlockCount(5000)
        self.registro.setFont(QFont("Consolas", 9))
        baixo = QWidget()
        lb = QVBoxLayout(baixo)
        lb.setContentsMargins(8, 4, 8, 8)
        lb.addWidget(self.pergunta)
        lb.addWidget(secundario(QLabel("Registro")))
        lb.addWidget(self.registro)

        divisor = QSplitter(Qt.Orientation.Vertical)
        divisor.addWidget(self.pilha)
        divisor.addWidget(baixo)
        divisor.setStretchFactor(0, 3)
        divisor.setStretchFactor(1, 1)
        divisor.setSizes([560, 220])

        central = QWidget()
        h = QHBoxLayout(central)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        h.addWidget(self.nav)
        h.addWidget(divisor, 1)
        self.setCentralWidget(central)

        # barra de status: tarefa em andamento + cancelar
        self.st_tarefa = QLabel("")
        self.st_prog = QProgressBar()
        self.st_prog.setRange(0, 0)
        self.st_prog.setMaximumWidth(140)
        self.st_cancelar = botao("Cancelar tarefa", self.cancelar)
        for w in (self.st_tarefa, self.st_prog, self.st_cancelar):
            self.statusBar().addPermanentWidget(w)
        self._ocupado(False)

        self.nav.setCurrentRow(0)
        erro = nucleo.erro_no_config()
        if erro:
            self.log(f"[AVISO] config.json com erro, usando o padrão: {erro}")

    # ---- navegação
    def _trocar(self, i):
        self.pilha.setCurrentIndex(i)
        self.pilha.currentWidget().ao_mostrar()

    def ir(self, chave, aba=None):
        self.nav.setCurrentRow([c for c, *_ in self.PAGINAS].index(chave))
        if aba and chave == "config":
            self.paginas["config"].abrir_aba(aba)

    # ---- registro
    def log(self, t):
        self.registro.appendPlainText(f"{datetime.now():%H:%M:%S}  {t}")

    # ---- tarefas em segundo plano
    def _ocupado(self, sim, titulo=""):
        self.st_tarefa.setText(f"Em andamento: {titulo}" if sim else "")
        self.st_prog.setVisible(sim)
        self.st_cancelar.setVisible(sim)

    def executar(self, titulo, funcao, ao_terminar=None):
        """Roda funcao(ui) em segundo plano; ao_terminar(resultado) na tela."""
        if self.tarefa:
            QMessageBox.information(self, "Aguarde",
                                    f"Já há uma tarefa em andamento: {self.tarefa}.")
            return
        ponte = PonteUI()
        ponte.sig_log.connect(self.log)
        ponte.sig_pergunta.connect(self._pergunta)
        ponte.sig_arquivo.connect(self._arquivo)
        ponte.sig_fim.connect(lambda r: self._fim(titulo, r, ao_terminar))
        self.ponte, self.tarefa = ponte, titulo
        self._ocupado(True, titulo)
        self.log(f"===== {titulo} =====")

        def alvo():
            try:
                ponte.sig_fim.emit(("ok", funcao(ponte)))
            except nucleo.Cancelado:
                ponte.sig_fim.emit(("cancelado", None))
            except Exception as e:
                ponte.sig_fim.emit(("erro", e, traceback.format_exc()))
        threading.Thread(target=alvo, daemon=True).start()

    def _pergunta(self, msg, opcoes):
        if self.ponte:
            self.pergunta.mostrar(msg, opcoes, self.ponte.responder)
            self.raise_()
            self.activateWindow()
            QApplication.alert(self)

    def _arquivo(self, msg):
        if not self.ponte:
            return
        caminho, _ = QFileDialog.getOpenFileName(
            self, msg, str(Path.home() / "Downloads"),
            "Planilhas (*.xlsx *.xls *.csv *.xlsm);;Todos (*.*)")
        self.ponte.responder(caminho or "")

    def _fim(self, titulo, r, ao_terminar):
        self.ponte, self.tarefa = None, None
        self.pergunta.hide()
        self._ocupado(False)
        if r[0] == "ok":
            self.log(f"{titulo}: concluído.")
            if ao_terminar:
                ao_terminar(r[1])
        elif r[0] == "cancelado":
            self.log(f"{titulo}: cancelado.")
        else:
            erro, detalhes = r[1], r[2]
            conhecido = type(erro).__name__ in ("ErroEntrada", "RuntimeError")
            self.log(f"{titulo}: ERRO — {erro}")
            if not conhecido:
                self.log(detalhes)
            QMessageBox.warning(self, titulo,
                                str(erro) if conhecido else
                                f"Erro inesperado: {erro}\n\nOs detalhes estão no registro.")

    def cancelar(self):
        if self.ponte:
            self.log("Cancelando... (fecha o navegador assim que possível)")
            self.ponte.cancelar()
            self.pergunta.hide()

    def mostrar_resultado(self, r):
        res, arquivos = r
        self.paginas["resultados"].mostrar(res, arquivos)
        self.ir("resultados")

    def closeEvent(self, ev):
        if self.tarefa:
            if QMessageBox.question(
                    self, "Tarefa em andamento",
                    f"'{self.tarefa}' ainda está rodando. Cancelar e sair?") \
                    != QMessageBox.StandardButton.Yes:
                ev.ignore()
                return
            self.ponte.cancelar()
        ev.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Automação Ábaco + Excel")
    app.setStyleSheet(ESTILO)
    janela = Janela()
    janela.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
