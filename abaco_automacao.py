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
from datetime import date, datetime
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
        QScrollArea, QSpinBox, QDoubleSpinBox, QSplitter, QStackedWidget, QStyle, QTableView,
        QTableWidget, QTableWidgetItem, QTabWidget, QTreeWidget,
        QTreeWidgetItem, QVBoxLayout, QWidget)
except ImportError:
    print("A tela precisa do Qt6 (PySide6), que ainda não está instalado.\n"
          "Dê dois cliques no INICIAR.bat: ele instala tudo e abre o programa.")
    sys.exit(1)

import nucleo_abaco as nucleo

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
         "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

import tema

ESTILO = tema.estilo()

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


def botao(t, ao_clicar=None, primario=False, discreto=False):
    b = QPushButton(t)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    if primario:
        b.setObjectName("primario")
    elif discreto:
        b.setObjectName("discreto")
    if ao_clicar:
        b.clicked.connect(ao_clicar)
    return b


def cartao(titulo, descricao=None):
    """Quadro com título; devolve (quadro, layout interno)."""
    q = QFrame()
    q.setObjectName("cartao")
    lay = QVBoxLayout(q)
    lay.setContentsMargins(20, 16, 20, 18)
    lay.setSpacing(10)
    t = QLabel(titulo)
    t.setObjectName("tituloCartao")
    lay.addWidget(t)
    if descricao:
        lay.addWidget(texto(descricao, sec=True))
    return q, lay


class Passo(QWidget):
    """Um passo da trilha numerada (sequência real do mês): círculo com o
    número à esquerda, linha ligando ao próximo, conteúdo à direita."""

    def __init__(self, numero, titulo, descricao=None, ultimo=False, setor="beneficios"):
        super().__init__()
        self.numero, self.setor = numero, setor
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(16)
        trilha = QVBoxLayout()
        trilha.setSpacing(0)
        trilha.setContentsMargins(0, 0, 0, 0)
        self.circulo = QLabel(str(numero))
        self.circulo.setFixedSize(32, 32)
        self.circulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        trilha.addWidget(self.circulo, 0, Qt.AlignmentFlag.AlignHCenter)
        fio = QFrame()
        fio.setObjectName("linhaPasso")
        fio.setFixedWidth(2)
        fio.setVisible(not ultimo)
        trilha.addWidget(fio, 1, Qt.AlignmentFlag.AlignHCenter)
        h.addLayout(trilha)
        self.lay = QVBoxLayout()
        self.lay.setContentsMargins(0, 4, 0, 18)
        self.lay.setSpacing(8)
        t = QLabel(titulo)
        t.setObjectName("tituloCartao")
        self.lay.addWidget(t)
        if descricao:
            self.lay.addWidget(texto(descricao, sec=True))
        h.addLayout(self.lay, 1)
        self.estado("a fazer")

    def estado(self, estado):
        self.circulo.setText("✓" if estado == "feito" else str(self.numero))
        self.circulo.setStyleSheet(tema.estilo_numero_passo(estado, self.setor))


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
        self.view.verticalHeader().setVisible(False)
        self.view.verticalHeader().setDefaultSectionSize(34)
        self.view.setShowGrid(False)
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
        self.corpo.setSpacing(16)
        self.corpo.setContentsMargins(0, 12, 0, 0)
        conteudo = QWidget()
        lay = QVBoxLayout(conteudo)
        lay.setContentsMargins(34, 22, 34, 22)
        lay.setSpacing(6)
        if self.titulo:
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
    titulo = "Visão geral"
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
    titulo = "Passos do mês"
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
        self.mes.currentIndexChanged.connect(self._mes_grande)
        self.ano.valueChanged.connect(self._mes_grande)

        # ---- competência em destaque (o mês é o objeto do trabalho)
        hero = QFrame()
        hero.setObjectName("cartao")
        hl = QHBoxLayout(hero)
        hl.setContentsMargins(22, 14, 22, 14)
        hl.setSpacing(18)
        esq = QVBoxLayout()
        esq.setSpacing(2)
        r = QLabel("Competência (mês do benefício)")
        r.setObjectName("rotuloPequeno")
        self.mes_grande = QLabel("")
        self.mes_grande.setObjectName("mesGrande")
        self.mes_grande.setTextFormat(Qt.TextFormat.RichText)
        esq.addWidget(r)
        esq.addWidget(self.mes_grande)
        hl.addLayout(esq)
        hl.addSpacing(10)
        hl.addWidget(self.mes, 0, Qt.AlignmentFlag.AlignVCenter)
        hl.addWidget(self.ano, 0, Qt.AlignmentFlag.AlignVCenter)
        hl.addStretch(1)
        dir_ = QVBoxLayout()
        dir_.setSpacing(2)
        r2 = QLabel("Dias úteis")
        r2.setObjectName("rotuloPequeno")
        r2.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.dias = QSpinBox()
        self.dias.setRange(0, 31)
        self.dias.setFixedWidth(84)
        self.dias.setToolTip("Calculado pelo calendário; ajuste se o decreto mudar.")
        dir_.addWidget(r2)
        dir_.addWidget(self.dias, 0, Qt.AlignmentFlag.AlignRight)
        hl.addLayout(dir_)
        self.corpo.addWidget(hero)
        self.info_dias = texto("", sec=True)
        self.corpo.addWidget(self.info_dias)

        # ---- os três passos (sequência real)
        self.p1 = Passo(1, "Coletar relatórios",
                        "Planilha base do mês anterior (aba CARGA, com o tipo de cada "
                        "servidor), relatório do SIGA, férias e faltas da Ábaco em "
                        "Excel e a planilha de plantões da Saúde. A partir do 2º mês o "
                        "programa usa a planilha que ele mesmo gerou.")
        self.arquivos = QTreeWidget()
        self.arquivos.setHeaderLabels(["Arquivo", "Origem", "Situação"])
        self.arquivos.setRootIsDecorated(False)
        self.arquivos.setMinimumHeight(220)
        self.arquivos.setAlternatingRowColors(False)
        self.arquivos.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.arquivos.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.arquivos.header().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.arquivos.header().setStretchLastSection(False)
        self.p1.lay.addWidget(self.arquivos)
        self.p1.lay.addWidget(linha(
            botao("Coletar relatórios", self.coletar, True),
            botao("Adicionar arquivo", self.adicionar_arquivo, discreto=True),
            botao("Abrir pasta de entrada", self.abrir_entrada, discreto=True),
            botao("Gerar modelos vazios", self.gerar_modelos, discreto=True)))
        self.corpo.addWidget(self.p1)

        self.p2 = Passo(2, "Conferir novos admitidos",
                        "Admissão, cargo, jornada e secretaria da Ábaco prevalecem sobre "
                        "o SIGA no cálculo; cada diferença vira pendência. A conferência "
                        "é feita na Consulta de Admissões.")
        self.aviso_robo = texto("", sec=True)
        self.bt_conferir = botao("Conferir pelo robô antigo", self.conferir, discreto=True)
        self.p2.lay.addWidget(linha(
            botao("Abrir Consulta de Admissões", lambda: self.j.ir("admissoes"), True),
            self.bt_conferir,
            botao("Configurar o robô", lambda: self.j.ir("config", "robo"), discreto=True)))
        self.p2.lay.addWidget(self.aviso_robo)
        self.corpo.addWidget(self.p2)

        self.p3 = Passo(3, "Calcular VR/VA",
                        "VR = (dias úteis − férias − faltas) × valor do tipo; plantonista "
                        "= plantões × valor. Aplica afastamentos, exonerações, o limite "
                        "salarial do diferenciado e inclui os novos admitidos (com "
                        "complementar). Gera a planilha no formato da Divisão, os arquivos "
                        "de importação e um modelo de ajustes para as pendências.",
                        ultimo=True)
        self.p3.lay.addWidget(linha(botao("Calcular VR/VA", self.calcular, True)))
        self.corpo.addWidget(self.p3)
        self._mes_grande()

    def _mes_grande(self):
        self.mes_grande.setText(
            f"{MESES[self.mes.currentIndex()]} "
            f"<span style='color:{tema.TINTA3}; font-weight:400'>{self.ano.value()}</span>")

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
        faltando = 0
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
            it.setToolTip(2, situacao)
            cor = tema.OK if arq else (tema.TINTA3 if nome in nucleo.OPCIONAIS
                                       else tema.ATENCAO)
            it.setForeground(2, QColor(cor))
            self.arquivos.addTopLevelItem(it)
            if not arq and nome not in nucleo.OPCIONAIS and nome != "plantoes":
                faltando += 1
        self.p1.estado("feito" if faltando == 0 else "agora")
        self.p2.estado("agora" if faltando == 0 else "a fazer")

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
class PaginaAdmissoes(Pagina):
    titulo = ""
    descricao = ""
    expandir = True

    def __init__(self, janela):
        super().__init__(janela)
        self.ultimo = None
        ano, mes = nucleo.competencia_padrao()
        self.mes = QComboBox()
        self.mes.addItems(MESES)
        self.mes.setCurrentIndex(mes - 1)
        self.ano = QSpinBox()
        self.ano.setRange(2020, 2100)
        self.ano.setValue(ano)
        self.mes.setToolTip("Competência: dias úteis da aba CARGA e planilha base")
        self.ano.setToolTip("Competência: dias úteis da aba CARGA e planilha base")

        # ---- período do SIGA em destaque (do 1º do mês até hoje)
        hero = QFrame()
        hero.setObjectName("destaque")
        hl = QHBoxLayout(hero)
        hl.setContentsMargins(22, 16, 22, 16)
        hl.setSpacing(12)
        esq = QVBoxLayout()
        esq.setSpacing(2)
        r = QLabel("Admitidos no SIGA, do início do mês até hoje")
        r.setObjectName("rotuloPequeno")
        hoje = date.today()
        self.periodo = QLabel(f"{hoje.replace(day=1):%d/%m/%Y}  →  {hoje:%d/%m/%Y}")
        self.periodo.setObjectName("periodoGrande")
        esq.addWidget(r)
        esq.addWidget(self.periodo)
        hl.addLayout(esq)
        hl.addStretch(1)
        rc = QLabel("Competência")
        rc.setObjectName("rotuloPequeno")
        hl.addWidget(rc)
        hl.addWidget(self.mes)
        hl.addWidget(self.ano)
        hl.addSpacing(10)
        hl.addWidget(botao("Gravar login do SIGA", self.gravar_login_siga))
        hl.addWidget(botao("Puxar matrículas do SIGA", self.do_siga, True))
        self.corpo.addWidget(hero)

        duas = QHBoxLayout()
        duas.setSpacing(18)
        # ---- coluna da esquerda: a lista e a ação principal
        q, lay = cartao("Matrículas a consultar")
        q.setFixedWidth(430)
        self.lista = QPlainTextEdit()
        self.lista.setObjectName("listaMatriculas")
        self.lista.setPlaceholderText("Uma por linha. Ex.:\n22771\n22769;2   (só o contrato 2)")
        self.lista.setMinimumHeight(150)
        self.lista.textChanged.connect(self._contar)
        lay.addWidget(self.lista, 1)
        self.contagem = texto("Nenhuma matrícula.", sec=True)
        lay.addWidget(self.contagem)
        lay.addWidget(linha(
            botao("Relatório da entrada", self.novos, discreto=True),
            botao("De planilha...", self.da_planilha, discreto=True),
            botao("Limpar", self.lista.clear, discreto=True)))
        div = QFrame()
        div.setObjectName("divisorFino")
        lay.addWidget(div)
        self.usar_conf = QCheckBox("Usar no cálculo do VR")
        self.usar_conf.setToolTip("Grava entrada\\conferencia_abaco.xlsx: admissão, cargo, "
                                  "jornada e secretaria da Ábaco prevalecem sobre o SIGA.")
        self.usar_conf.setChecked(True)
        self.devagar = QCheckBox("Devagar, para acompanhar no navegador")
        self.devagar.setChecked(bool(self.j.cfg.get("reproduzir_devagar", False)))
        lay.addWidget(self.usar_conf)
        lay.addWidget(self.devagar)
        principal = botao("Consultar e extrair para Excel", self.consultar, True)
        principal.setProperty("grande", True)
        principal.setMinimumHeight(40)
        lay.addWidget(principal)
        duas.addWidget(q)

        # ---- coluna da direita: o resultado
        q, lay = cartao("Resultado")
        self.filtro = QLineEdit()
        self.filtro.setPlaceholderText("Filtrar por nome ou matrícula")
        self.bt_abrir = botao("Abrir o Excel", self.abrir_excel)
        self.bt_abrir.setEnabled(False)
        lay.addWidget(linha(self.filtro, self.bt_abrir,
                            botao("Abrir pasta das consultas", self.abrir_pasta,
                                  discreto=True), esticar=False))
        self.tabela = TabelaDados()
        self.tabela.setMinimumHeight(240)
        self.filtro.textChanged.connect(self.tabela.filtrar)
        lay.addWidget(self.tabela, 1)
        lay.addWidget(texto("O Excel abre na aba CARGA, no formato da planilha da "
                            "Divisão. As abas Consultas e Fichas vêm depois. Só "
                            "consulta: não altera nada no sistema.", sec=True))
        duas.addWidget(q, 1)
        self.corpo.addLayout(duas, 1)

    def competencia(self):
        return self.ano.value(), self.mes.currentIndex() + 1

    def _contar(self):
        n = len(nucleo.ler_lista_matriculas(self.lista.toPlainText()))
        self.contagem.setText(f"{n} matrícula(s)." if n else "Nenhuma matrícula.")

    def do_siga(self):
        ano, mes = self.competencia()
        cfg = self.j.cfg

        def fim(r):
            if not r:
                return
            txt, resumo = r
            if not txt:
                QMessageBox.information(self, "SIGA", resumo + "\n\nNenhuma "
                                        "matrícula nova para consultar.")
                return
            self.lista.setPlainText(txt)
            QMessageBox.information(self, "SIGA", resumo + "\n\nA lista de "
                                    "matrículas também está pronta: confira e "
                                    "clique em Consultar e extrair para Excel.")
        self.j.executar("Matrículas do SIGA",
                        lambda ui: nucleo.matriculas_do_siga(cfg, ui, ano, mes), fim)

    def gravar_login_siga(self):
        cfg = self.j.cfg
        QMessageBox.information(
            self, "Gravar login do SIGA",
            "Vai abrir o navegador na tela de login do SIGA. Faça SÓ o login "
            "(usuário, senha, Entrar) e FECHE o navegador.\n\nO usuário e a senha "
            f"ficam gravados em {cfg.get('arquivo_fluxo_siga') or 'fluxo_siga.py'}, "
            "na pasta do programa, para o robô entrar sozinho.")
        self.j.executar("Gravar login do SIGA",
                        lambda ui: nucleo.gravar(cfg, ui,
                                                 cfg.get("arquivo_fluxo_siga") or "fluxo_siga.py",
                                                 cfg.get("url_siga")))

    def novos(self):
        ano, mes = self.competencia()
        cfg = self.j.cfg

        def fim(r):
            txt, n = r
            if not n:
                QMessageBox.information(self, "Novos admitidos",
                                        "Nenhum novo admitido para essa competência "
                                        "no relatório do SIGA.")
                return
            self.lista.setPlainText(txt)
            self.j.log(f"{n} novo(s) admitido(s) na lista.")
        self.j.executar("Novos admitidos",
                        lambda ui: nucleo.novos_para_consultar(cfg, ano, mes), fim)

    def da_planilha(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Planilha com as matrículas", str(nucleo.PASTA),
            "Planilhas (*.xlsx *.xls *.csv *.xlsm)")
        if not caminho:
            return
        try:
            txt = nucleo.matriculas_da_planilha(caminho)
        except Exception as e:
            QMessageBox.warning(self, "Carregar de planilha", str(e))
            return
        self.lista.setPlainText(txt)

    def consultar(self):
        itens = nucleo.ler_lista_matriculas(self.lista.toPlainText())
        if not itens:
            QMessageBox.information(self, "Consulta de admissões",
                                    "Informe pelo menos uma matrícula.")
            return
        ano, mes = self.competencia()
        cfg, conf, devagar = self.j.cfg, self.usar_conf.isChecked(), \
            self.devagar.isChecked()

        def fim(r):
            if not r:
                return
            self.ultimo = r["excel"]
            self.bt_abrir.setEnabled(True)
            import pandas as pd
            df = pd.DataFrame(r["registros"])
            cols = [c for c in ["Matrícula", "Cnt", "Nome", "Cargo", "Admissão",
                                "Jornada", "Horas semanais", "Tipo sugerido",
                                "Lotação/Secretaria", "Erro"] if c in df.columns]
            self.tabela.definir(df[cols])
            erros = int((df.get("Erro", pd.Series(dtype=str)).fillna("") != "").sum())
            msg = (f"{len(df)} linha(s) lida(s), {erros} com erro.\n\n"
                   f"Excel: {r['excel'].name}")
            if r.get("conferencia"):
                msg += ("\n\nTambém gravei entrada\\conferencia_abaco.xlsx: o "
                        "cálculo do VR vai usar esses dados.")
            if QMessageBox.question(self, "Consulta pronta", msg + "\n\nAbrir o Excel?") \
                    == QMessageBox.StandardButton.Yes:
                nucleo.abrir_no_sistema(r["excel"])
        self.j.executar(
            "Consulta de admissões",
            lambda ui: nucleo.consultar_admissoes(cfg, ui, itens, ano, mes, conf, devagar),
            fim)

    def abrir_excel(self):
        if self.ultimo and Path(self.ultimo).exists():
            nucleo.abrir_no_sistema(self.ultimo)

    def abrir_pasta(self):
        pasta = nucleo.PASTA / self.j.cfg["pasta_saida"] / "consultas"
        pasta.mkdir(parents=True, exist_ok=True)
        nucleo.abrir_no_sistema(pasta)


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
            "grave a pesquisa de uma matrícula na tela de consulta do servidor. "
            "Para testar uma gravação, use Reproduzir fluxo: o robô entra no "
            "sistema e repete os passos (no sistema real: só consultas e "
            "relatórios). Se precisar procurar algum item, ele grava um "
            "fluxo _corrigido e, nas próximas vezes, vai direto.")
        self.devagar = QCheckBox("Reproduzir devagar (para acompanhar)")
        self.devagar.setChecked(bool(self.j.cfg.get("reproduzir_devagar", False)))
        self.devagar.toggled.connect(self.trocar_velocidade)
        lay.addWidget(linha(botao("Gravar fluxo", self.gravar, True),
                            botao("Reproduzir fluxo...", self.reproduzir),
                            botao("Ver o código do fluxo", self.abrir_fluxo)))
        lay.addWidget(linha(self.devagar))
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

    def trocar_velocidade(self, marcado):
        self.j.cfg["reproduzir_devagar"] = bool(marcado)
        try:
            nucleo.salvar_config(self.j.cfg)
        except Exception:
            pass

    def reproduzir(self):
        cfg = self.j.cfg
        padrao = nucleo.PASTA / cfg["arquivo_fluxo"]
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Qual fluxo reproduzir?",
            str(padrao if padrao.exists() else nucleo.PASTA),
            "Fluxos gravados (*.py)")
        if not caminho:
            return

        def fim(r):
            if not r:
                return
            if r["salvos"]:
                self.j.log(f"{len(r['salvos'])} arquivo(s) baixado(s) em "
                           f"{r['salvos'][0].parent}")
            extra = ""
            if r.get("corrigido"):
                extra = (f"\n\nO robô aprendeu o caminho e salvou "
                         f"{r['corrigido'].name}. Nas próximas vezes ele usa esse "
                         f"arquivo sozinho e vai direto.")
            if r["parou"]:
                i, passo, motivo = r["parou"]
                QMessageBox.warning(self, "O fluxo parou",
                                    f"Parou no passo {i} de {r['passos']}:\n{passo}\n\n"
                                    f"Motivo: {motivo}\n\nSe a tela do sistema mudou, "
                                    f"grave o fluxo de novo." + extra)
            else:
                QMessageBox.information(self, "Fluxo reproduzido",
                                        f"Os {r['passos']} passos foram repetidos "
                                        f"até o fim." + extra)
        self.j.executar("Reproduzir fluxo",
                        lambda ui: nucleo.reproduzir_fluxo(cfg, ui, caminho), fim)

    def abrir_fluxo(self):
        f = nucleo.PASTA / self.j.cfg["arquivo_fluxo"]
        if not f.exists():
            QMessageBox.information(self, "Sem fluxo", "Nenhum fluxo gravado ainda.")
        elif sys.platform == "win32":
            # abrir no Bloco de Notas: com dois cliques o Windows EXECUTARIA o .py
            import subprocess
            subprocess.Popen(["notepad.exe", str(f)])
        else:
            nucleo.abrir_no_sistema(f)


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
        # SIGA: login e a planilha dos admitidos (como no programa do Gilson)
        self.siga_usuario = QLineEdit()
        self.siga_senha = QLineEdit()
        self.siga_senha.setEchoMode(QLineEdit.EchoMode.Password)
        self.siga_usuario.setPlaceholderText("vazio: usa o login gravado ou você entra à mão")
        self.planilha_adm = QLineEdit()
        self.adm_qtde = QSpinBox()
        self.adm_qtde.setRange(0, 31)
        self.adm_valor = QDoubleSpinBox()
        self.adm_valor.setRange(0, 10000)
        self.adm_valor.setDecimals(2)
        self.abrir_adm = QCheckBox("Abrir a planilha no Excel ao terminar")
        qv = QHBoxLayout()
        qv.addWidget(QLabel("QTDE"))
        qv.addWidget(self.adm_qtde)
        qv.addSpacing(16)
        qv.addWidget(QLabel("VALOR (R$)"))
        qv.addWidget(self.adm_valor)
        qv.addStretch(1)
        f.addRow("Usuário do SIGA:", self.siga_usuario)
        f.addRow("Senha do SIGA:", self.siga_senha)
        f.addRow("Planilha dos admitidos (SIGA):", self.planilha_adm)
        f.addRow("Cada admitido entra com:", qv)
        f.addRow("", self.abrir_adm)
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
        self.siga_usuario.setText(str(cfg.get("siga_usuario") or ""))
        self.siga_senha.setText(str(cfg.get("siga_senha") or ""))
        self.planilha_adm.setText(cfg.get("planilha_admissoes_siga") or "admissoes.xlsx")
        self.adm_qtde.setValue(int(cfg.get("admissoes_qtde", 19) or 0))
        self.adm_valor.setValue(float(cfg.get("admissoes_valor", 56) or 0))
        self.abrir_adm.setChecked(bool(cfg.get("abrir_planilha_admissoes", True)))

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
            "siga_usuario": self.siga_usuario.text().strip(),
            "siga_senha": self.siga_senha.text(),
            "planilha_admissoes_siga": self.planilha_adm.text().strip() or "admissoes.xlsx",
            "admissoes_qtde": self.adm_qtde.value(),
            "admissoes_valor": (int(self.adm_valor.value())
                                if self.adm_valor.value().is_integer()
                                else self.adm_valor.value()),
            "abrir_planilha_admissoes": self.abrir_adm.isChecked(),
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
    descricao_exe = ("Programa empacotado (AutomacaoAbaco.exe, arquivo único): o "
                     "Python, as bibliotecas e o config.json padrão vêm dentro "
                     "dele, sem instalar nada. Aqui dá "
                     "para conferir e baixar o navegador controlado (Chromium); "
                     "sem ele, o robô usa o Edge do Windows.")

    def __init__(self, janela):
        if nucleo.EMPACOTADO:   # antes de montar a tela (a classe Pagina lê isto)
            PaginaInstalacao.descricao = PaginaInstalacao.descricao_exe
        super().__init__(janela)
        q, lay = cartao("Situação")
        self.lista = QTreeWidget()
        self.lista.setHeaderLabels(["Item", "Situação"])
        self.lista.setRootIsDecorated(False)
        self.lista.setMinimumHeight(220)
        self.lista.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(self.lista)
        lay.addWidget(linha(botao("Baixar o navegador controlado (Chromium)"
                                  if nucleo.EMPACOTADO else
                                  "Instalar / atualizar bibliotecas e navegador",
                                  self.instalar, True),
                            botao("Verificar de novo", self.ao_mostrar)))
        self.corpo.addWidget(q)

    def ao_mostrar(self):
        self.lista.clear()
        self.lista.addTopLevelItem(QTreeWidgetItem(
            ["Python", f"{sys.version.split()[0]} — "
                       + ("embutido no programa" if nucleo.EMPACOTADO
                          else sys.executable)]))
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
# ==========================================================================
# Tela inicial: os dois setores (e a engrenagem com o resto)
# ==========================================================================
SETORES = {
    "beneficios": ("Benefícios VR/VA", ["beneficios", "resultados", "inicio"]),
    "admissoes": ("Consulta de Admissões", ["admissoes"]),
    # a engrenagem: todas as funções do programa, como na versão antiga
    "dev": ("Modo desenvolvedor", ["beneficios", "resultados", "inicio", "admissoes",
                                   "config", "instalacao", "diagnostico"]),
}

# barra lateral: grupos (só aparecem no modo desenvolvedor) e páginas
NAV = [("grupo", "Benefícios VR/VA"), "beneficios", "resultados", "inicio",
       ("grupo", "Consulta de Admissões"), "admissoes",
       ("grupo", "Ferramentas"), "config", "instalacao", "diagnostico"]

ICONE_BENEFICIOS = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96">
<g fill="none" stroke="#ffffff" stroke-width="4" stroke-linejoin="round" stroke-linecap="round">
<rect x="30" y="14" width="56" height="38" rx="6" transform="rotate(-14 58 33)" fill="#2b74b5"/>
<rect x="10" y="30" width="64" height="44" rx="6" fill="#2b74b5"/>
<line x1="10" y1="42" x2="74" y2="42"/>
<line x1="20" y1="56" x2="42" y2="56"/>
<line x1="20" y1="64" x2="34" y2="64"/>
<rect x="52" y="54" width="14" height="10" rx="5"/></g></svg>"""

ICONE_ADMISSOES = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96">
<path d="M22 8h36l16 16v60a4 4 0 0 1-4 4H22a4 4 0 0 1-4-4V12a4 4 0 0 1 4-4z"
 fill="#f7f4f0" stroke="#3a4656" stroke-width="3.5" stroke-linejoin="round"/>
<path d="M58 8v16h16" fill="none" stroke="#3a4656" stroke-width="3.5" stroke-linejoin="round"/>
<circle cx="34" cy="26" r="6" fill="#e58a6a" stroke="#3a4656" stroke-width="2.5"/>
<path d="M24 42c1-7 5-10 10-10s9 3 10 10z" fill="#e58a6a" stroke="#3a4656" stroke-width="2.5"/>
<g stroke="#3a4656" stroke-width="3" stroke-linecap="round">
<line x1="26" y1="54" x2="44" y2="54"/><line x1="26" y1="62" x2="40" y2="62"/>
<line x1="26" y1="70" x2="42" y2="70"/></g>
<circle cx="60" cy="58" r="15" fill="#d6e6f5" stroke="#3a4656" stroke-width="3.5"/>
<circle cx="60" cy="53" r="4.5" fill="#5b8fc7" stroke="#3a4656" stroke-width="2"/>
<path d="M52 66c1-5 4-7 8-7s7 2 8 7" fill="#5b8fc7" stroke="#3a4656" stroke-width="2"/>
<line x1="71" y1="69" x2="84" y2="82" stroke="#3a4656" stroke-width="6" stroke-linecap="round"/>
<line x1="71" y1="69" x2="84" y2="82" stroke="#5b8fc7" stroke-width="2.5" stroke-linecap="round"/>
</svg>"""

ICONE_ENGRENAGEM = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
<g fill="none" stroke="#3a4656" stroke-width="2.6" stroke-linejoin="round">
<path d="M20.5 5h7l1.2 5.2 3.3 1.9 5.1-1.6 3.5 6.1-3.9 3.6v3.8l3.9 3.6-3.5 6.1-5.1-1.6-3.3 1.9L27.5 43h-7l-1.2-5.2-3.3-1.9-5.1 1.6-3.5-6.1 3.9-3.6v-3.8L7.4 16.6l3.5-6.1 5.1 1.6 3.3-1.9z"/>
<circle cx="24" cy="24" r="6.5"/></g></svg>"""


def icone_svg(svg, lado):
    """QIcon a partir de um SVG (nítido em telas com zoom)."""
    from PySide6.QtCore import QByteArray, QSize
    from PySide6.QtGui import QIcon, QPainter, QPixmap
    from PySide6.QtSvg import QSvgRenderer
    escala = 2
    pix = QPixmap(QSize(lado * escala, lado * escala))
    pix.fill(Qt.GlobalColor.transparent)
    pintor = QPainter(pix)
    QSvgRenderer(QByteArray(svg.encode("utf-8"))).render(pintor)
    pintor.end()
    pix.setDevicePixelRatio(escala)
    return QIcon(pix)


class CartaoSetor(QPushButton):
    """Botão grande da tela inicial: ícone em cima, nome embaixo."""

    def __init__(self, nome, svg, objeto, ao_clicar, desc=""):
        super().__init__()
        from PySide6.QtCore import QSize
        self.setObjectName(objeto)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(360, 262)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 34, 20, 30)
        lay.setSpacing(14)
        ic = QLabel()
        ic.setPixmap(icone_svg(svg, 92).pixmap(QSize(92, 92)))
        ic.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ic.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        rot = QLabel(nome)
        rot.setObjectName(objeto + "Nome")
        rot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rot.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        lay.addWidget(ic, 1)
        lay.addWidget(rot)
        if desc:
            d = QLabel(desc)
            d.setObjectName(objeto + "Desc")
            d.setAlignment(Qt.AlignmentFlag.AlignCenter)
            d.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            lay.addWidget(d)
        self.clicked.connect(ao_clicar)


class TelaInicial(QWidget):
    def __init__(self, janela):
        super().__init__()
        from PySide6.QtCore import QSize
        from PySide6.QtWidgets import QGraphicsDropShadowEffect, QToolButton
        self.setObjectName("telaInicial")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        grade = QGridLayout(self)
        grade.setContentsMargins(36, 24, 36, 36)
        cartoes = QHBoxLayout()
        cartoes.setSpacing(24)
        cartoes.addStretch(1)
        cartoes.setSpacing(28)
        cartoes.addWidget(CartaoSetor("Benefícios VR/VA", ICONE_BENEFICIOS, "cartaoAzul",
                                      lambda: janela.entrar("beneficios"),
                                      "Carga mensal do vale-refeição e do vale-alimentação"))
        cartoes.addWidget(CartaoSetor("Consulta de Admissões", ICONE_ADMISSOES,
                                      "cartaoBege", lambda: janela.entrar("admissoes"),
                                      "Novos admitidos do SIGA e fichas do e-Turmalina"))
        cartoes.addStretch(1)
        grade.setRowStretch(0, 2)
        saud = QLabel("Prefeitura de Nova Lima, Divisão de Benefícios")
        saud.setObjectName("saudacao")
        saud.setAlignment(Qt.AlignmentFlag.AlignCenter)
        grade.addWidget(saud, 1, 0, 1, 3)
        grade.setRowMinimumHeight(2, 24)
        grade.addLayout(cartoes, 3, 0, 1, 3)
        grade.setRowStretch(4, 3)
        dica = QLabel("A engrenagem abre o modo desenvolvedor, com todas as funções")
        dica.setObjectName("dicaInicial")
        grade.addWidget(dica, 5, 2, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom)
        eng = QToolButton()
        eng.setObjectName("engrenagem")
        eng.setIcon(icone_svg(ICONE_ENGRENAGEM, 30))
        eng.setIconSize(QSize(30, 30))
        eng.setFixedSize(64, 64)
        eng.setCursor(Qt.CursorShape.PointingHandCursor)
        eng.setToolTip("Modo desenvolvedor: todas as funções do programa")
        sombra = QGraphicsDropShadowEffect(eng)
        sombra.setBlurRadius(18)
        sombra.setOffset(0, 3)
        sombra.setColor(QColor(0, 0, 0, 60))
        eng.setGraphicsEffect(sombra)
        eng.clicked.connect(lambda: janela.entrar("dev"))
        grade.addWidget(eng, 5, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom)


class Janela(QMainWindow):
    PAGINAS = [("beneficios", "Passos do mês", QStyle.StandardPixmap.SP_FileDialogDetailedView,
                PaginaBeneficios),
               ("admissoes", "Consulta de admissões", QStyle.StandardPixmap.SP_DirOpenIcon,
                PaginaAdmissoes),
               ("resultados", "Resultados", QStyle.StandardPixmap.SP_FileDialogContentsView,
                PaginaResultados),
               ("inicio", "Visão geral", QStyle.StandardPixmap.SP_DirHomeIcon, PaginaInicio),
               ("config", "Configurações", QStyle.StandardPixmap.SP_FileDialogListView,
                PaginaConfig),
               ("instalacao", "Instalação", QStyle.StandardPixmap.SP_DriveHDIcon,
                PaginaInstalacao),
               ("diagnostico", "Diagnóstico", QStyle.StandardPixmap.SP_FileDialogInfoView,
                PaginaDiagnostico)]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automação - Fluxos RH")
        self.resize(1280, 820)
        self.setMinimumSize(1080, 700)
        self.cfg = nucleo.carregar_config()
        self.ponte = None
        self.tarefa = None

        self.nav = QListWidget()
        self.nav.setObjectName("navegacao")
        self.nav.setFixedWidth(220)
        self.nav.setCursor(Qt.CursorShape.PointingHandCursor)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.pilha = QStackedWidget()
        self.paginas = {}
        rotulos = {}
        for chave, rotulo, icone, classe in self.PAGINAS:
            rotulos[chave] = rotulo
            pagina = classe(self)
            self.paginas[chave] = pagina
            self.pilha.addWidget(pagina)
        self.linhas_nav = []
        for item in NAV:
            if isinstance(item, tuple):             # título de grupo
                it = QListWidgetItem(item[1])
                it.setFlags(Qt.ItemFlag.NoItemFlags)
                f = it.font()
                f.setPointSizeF(8.5)
                f.setBold(True)
                it.setFont(f)
                self.nav.addItem(it)
                self.linhas_nav.append(None)
            else:
                self.nav.addItem(QListWidgetItem(rotulos[item]))
                self.linhas_nav.append(item)
        self.nav.currentRowChanged.connect(self._trocar)

        # registro + painel de pergunta
        self.pergunta = PainelPergunta()
        self.registro = QPlainTextEdit()
        self.registro.setReadOnly(True)
        self.registro.setMaximumBlockCount(5000)
        self.registro.setObjectName("registro")
        self.registro.setFont(QFont(tema.FAMILIA_MONO, 9))
        baixo = QFrame()
        baixo.setObjectName("caixaRegistro")
        lb = QVBoxLayout(baixo)
        lb.setContentsMargins(0, 0, 0, 0)
        lb.setSpacing(0)
        cab = QHBoxLayout()
        cab.setContentsMargins(16, 6, 10, 4)
        cab.setSpacing(10)
        tr = QLabel("Registro")
        tr.setObjectName("tituloRegistro")
        cab.addWidget(tr)
        cab.addWidget(texto("o que o programa está fazendo", wrap=False, sec=True))
        cab.addStretch(1)
        self.bt_recolher = botao("Recolher", self._recolher_registro, discreto=True)
        cab.addWidget(self.bt_recolher)
        lb.addLayout(cab)
        caixa_p = QVBoxLayout()
        caixa_p.setContentsMargins(16, 0, 16, 6)
        caixa_p.addWidget(self.pergunta)
        lb.addLayout(caixa_p)
        lb.addWidget(self.registro, 1)
        self.caixa_registro = baixo

        divisor = QSplitter(Qt.Orientation.Vertical)
        divisor.setHandleWidth(1)
        self.divisor = divisor
        divisor.addWidget(self.pilha)
        divisor.addWidget(baixo)
        divisor.setStretchFactor(0, 3)
        divisor.setStretchFactor(1, 1)
        divisor.setSizes([650, 130])

        area = QWidget()
        self.area = area
        va = QVBoxLayout(area)
        va.setContentsMargins(0, 0, 0, 0)
        va.setSpacing(0)
        topo = QFrame()
        topo.setObjectName("topoSetor")
        ht = QHBoxLayout(topo)
        ht.setContentsMargins(10, 6, 14, 6)
        topo.setFixedHeight(56)
        ht.setContentsMargins(14, 0, 20, 0)
        ht.setSpacing(14)
        self.bt_inicio = botao("←  Início", self.voltar_inicio)
        self.bt_inicio.setObjectName("voltar")
        sep = QFrame()
        sep.setObjectName("separadorTopo")
        sep.setFixedSize(1, 24)
        self.selo = QFrame()
        self.selo.setObjectName("selo")
        self.selo.setFixedSize(10, 28)
        self.titulo_setor = QLabel("")
        self.titulo_setor.setObjectName("tituloSetor")
        ht.addWidget(self.bt_inicio)
        ht.addWidget(sep)
        ht.addWidget(self.selo)
        ht.addWidget(self.titulo_setor)
        self.selo_dev = QLabel("todas as funções do programa")
        self.selo_dev.setObjectName("seloDev")
        self.selo_dev.setVisible(False)
        ht.addWidget(self.selo_dev, 0, Qt.AlignmentFlag.AlignVCenter)
        ht.addStretch(1)
        va.addWidget(topo)
        corpo = QWidget()
        h = QHBoxLayout(corpo)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        h.addWidget(self.nav)
        h.addWidget(divisor, 1)
        va.addWidget(corpo, 1)

        self.inicial = TelaInicial(self)
        self.telas = QStackedWidget()
        self.telas.addWidget(self.inicial)
        self.telas.addWidget(area)
        self.setCentralWidget(self.telas)
        self.setor = None

        # barra de status: tarefa em andamento + cancelar
        self.st_tarefa = QLabel("")
        self.st_prog = QProgressBar()
        self.st_prog.setRange(0, 0)
        self.st_prog.setMaximumWidth(140)
        self.st_cancelar = botao("Cancelar tarefa", self.cancelar)
        for w in (self.st_tarefa, self.st_prog, self.st_cancelar):
            self.statusBar().addPermanentWidget(w)
        self._ocupado(False)

        erro = nucleo.erro_no_config()
        if erro:
            self.log(f"[AVISO] config.json com erro, usando o padrão: {erro}")

    # ---- navegação: tela inicial > setor > páginas do setor
    def _trocar(self, i):
        if i < 0 or i >= len(self.linhas_nav) or self.linhas_nav[i] is None:
            return
        pagina = self.paginas[self.linhas_nav[i]]
        self.pilha.setCurrentWidget(pagina)
        pagina.ao_mostrar()

    def entrar(self, setor, chave=None):
        """Abre um setor mostrando só as páginas dele. No modo desenvolvedor
        (engrenagem), mostra todas, agrupadas."""
        titulo, chaves = SETORES[setor]
        self.setor = setor
        dev = setor == "dev"
        for i, c in enumerate(self.linhas_nav):
            self.nav.setRowHidden(i, (not dev) if c is None else c not in chaves)
        self.nav.setVisible(len(chaves) > 1)
        self.titulo_setor.setText(titulo)
        self.selo_dev.setVisible(dev)
        # cor do setor: azul (Benefícios), areia (Admissões), neutro (dev)
        self.area.setStyleSheet(tema.estilo(setor))
        self.telas.setCurrentIndex(1)
        alvo = self.linhas_nav.index(chave if chave in chaves else chaves[0])
        if self.nav.currentRow() == alvo:
            self._trocar(alvo)
        else:
            self.nav.setCurrentRow(alvo)

    def _recolher_registro(self):
        aberto = self.registro.isVisible()
        self.registro.setVisible(not aberto)
        self.bt_recolher.setText("Mostrar" if aberto else "Recolher")
        total = sum(self.divisor.sizes())
        self.divisor.setSizes([total - 40, 40] if aberto else [total - 130, 130])

    def voltar_inicio(self):
        self.telas.setCurrentIndex(0)

    def ir(self, chave, aba=None):
        if self.setor and chave in SETORES[self.setor][1] and self.telas.currentIndex() == 1:
            setor = self.setor                       # fica no setor atual
        else:
            setor = next(s for s, (_, cs) in SETORES.items() if chave in cs)
        self.entrar(setor, chave)
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
            # não puxa a janela para a frente (ela cobriria o navegador em que a
            # pessoa está trabalhando): só pisca na barra de tarefas
            if self.isMinimized():
                self.showNormal()
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
    if "--verificar" in sys.argv:          # confere o pacote sem abrir a tela
        ok, linhas = nucleo.verificar_instalacao()
        try:
            print("\n".join(linhas))
        except Exception:
            pass                            # .exe sem console: vale o arquivo
        sys.exit(0 if ok else 1)
    app = QApplication(sys.argv)
    app.setApplicationName("Automação - Fluxos RH")
    app.setStyle("Fusion")                 # desenho igual em todo Windows
    tema.carregar_fontes()
    app.setFont(QFont(tema.FAMILIA, 10))
    app.setStyleSheet(tema.estilo())
    janela = Janela()
    janela.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
