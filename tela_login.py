# -*- coding: utf-8 -*-
"""Tela de login e cadastro (KeyAuth) do Automação - Fluxos RH.

Aparece antes da janela principal. À esquerda, a marca do programa; à
direita, Entrar e Criar conta lado a lado num seletor. A conversa com o
KeyAuth roda numa thread, para a janela não travar.
"""
from PySide6.QtCore import QByteArray, QObject, QSettings, QSize, Qt, QThread, Signal, Slot
from PySide6.QtGui import QAction, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (QButtonGroup, QCheckBox, QDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QStackedWidget, QVBoxLayout, QWidget)

import autenticacao as auth
import tema

AZUL, AZUL_ESCURO = "#2B74B5", "#1F5A8F"

CARTAO = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96">
<g fill="none" stroke="#ffffff" stroke-width="3.2" stroke-linejoin="round" stroke-linecap="round">
<rect x="30" y="14" width="56" height="38" rx="6" transform="rotate(-14 58 33)" fill="#2c6fa9"/>
<rect x="10" y="30" width="64" height="44" rx="6" fill="#2a6aa3"/>
<line x1="10" y1="42" x2="74" y2="42"/><line x1="20" y1="56" x2="42" y2="56"/>
<line x1="20" y1="64" x2="34" y2="64"/><rect x="52" y="54" width="14" height="10" rx="5"/></g></svg>"""

LUPA = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96">
<g fill="none" stroke="#ffffff" stroke-width="3.2" stroke-linejoin="round" stroke-linecap="round">
<path d="M22 10h34l14 14v58a4 4 0 0 1-4 4H22a4 4 0 0 1-4-4V14a4 4 0 0 1 4-4z"/>
<path d="M56 10v14h14"/><line x1="28" y1="34" x2="48" y2="34"/><line x1="28" y1="44" x2="42" y2="44"/>
<circle cx="56" cy="60" r="13"/><line x1="66" y1="70" x2="80" y2="84" stroke-width="5"/></g></svg>"""

OLHO = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><g fill="none"
stroke="#8695A3" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></g></svg>"""

OLHO_FECHADO = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><g fill="none"
stroke="#8695A3" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
<path d="M3 3l18 18"/><path d="M10.6 5.1A10.8 10.8 0 0 1 12 5c6.4 0 10 7 10 7a17 17 0 0 1-3.2 4.1"/>
<path d="M6.6 6.6C3.8 8.4 2 12 2 12s3.6 7 10 7a9.7 9.7 0 0 0 5.4-1.6"/>
<path d="M9.9 9.9a3 3 0 0 0 4.2 4.2"/></g></svg>"""


def _icone(svg, lado):
    from PySide6.QtSvg import QSvgRenderer
    escala = 2
    pix = QPixmap(QSize(lado * escala, lado * escala))
    pix.fill(Qt.GlobalColor.transparent)
    p = QPainter(pix)
    QSvgRenderer(QByteArray(svg.encode("utf-8"))).render(p)
    p.end()
    pix.setDevicePixelRatio(escala)
    return QIcon(pix)


ESTILO = f"""
QDialog#telaLogin {{ background: {tema.SUPERFICIE}; }}
QWidget#pagina, QStackedWidget#pilhaLogin {{ background: {tema.SUPERFICIE}; }}
QWidget#pagina QLabel, QWidget#pagina QCheckBox {{ background: transparent; }}
QFrame#marca {{ background: qlineargradient(x1:0, y1:0, x2:0.35, y2:1,
    stop:0 {AZUL_ESCURO}, stop:1 {AZUL}); }}
QLabel#marcaOrgao {{ color: rgba(255,255,255,0.78); font-size: 13px; font-weight: 500; }}
QLabel#marcaNome {{ color: white; font-size: 40px; font-weight: 700; }}
QLabel#marcaSub {{ color: rgba(255,255,255,0.86); font-size: 15px; }}
QLabel#marcaItem {{ color: white; font-size: 14px; font-weight: 600; }}
QLabel#marcaItemSub {{ color: rgba(255,255,255,0.72); font-size: 12.5px; }}
QLabel#marcaRodape {{ color: rgba(255,255,255,0.6); font-size: 12px; }}
QFrame#linhaMarca {{ background: rgba(255,255,255,0.22); max-height: 1px; border: none; }}

QFrame#seletor {{ background: #EEF1F4; border-radius: 10px; }}
QPushButton#aba {{ background: transparent; border: none; border-radius: 8px;
    padding: 8px 0; color: {tema.TINTA2}; font-size: 13.5px; font-weight: 600; }}
QPushButton#aba:checked {{ background: white; color: {tema.TINTA}; }}
QPushButton#aba:hover:!checked {{ color: {tema.TINTA}; }}

QLabel#tituloForm {{ color: {tema.TINTA}; font-size: 24px; font-weight: 700; }}
QLabel#subForm {{ color: {tema.TINTA2}; font-size: 13.5px; }}
QLabel#rotulo {{ color: {tema.TINTA}; font-size: 12.5px; font-weight: 600; }}
QLabel#dica {{ color: {tema.TINTA3}; font-size: 12px; }}
QLineEdit#campo {{ background: white; border: 1px solid {tema.LINHA_CAMPO}; border-radius: 8px;
    padding: 0 12px; min-height: 40px; font-size: 14px; color: {tema.TINTA};
    selection-background-color: {AZUL}; }}
QLineEdit#campo:focus {{ border: 2px solid {AZUL}; padding: 0 11px; }}
QLineEdit#campo:disabled {{ background: {tema.PAPEL}; color: {tema.TINTA3}; }}
QCheckBox#lembrar {{ color: {tema.TINTA2}; font-size: 13px; spacing: 8px; }}
QPushButton#acao {{ background: {AZUL}; color: white; border: none; border-radius: 8px;
    min-height: 44px; font-size: 14.5px; font-weight: 600; }}
QPushButton#acao:hover {{ background: #3482C6; }}
QPushButton#acao:pressed {{ background: {AZUL_ESCURO}; }}
QPushButton#acao:disabled {{ background: #9DBEDD; }}
QPushButton#link {{ background: transparent; border: none; color: {AZUL}; font-size: 13px;
    font-weight: 600; padding: 2px 0; text-align: left; }}
QPushButton#link:hover {{ color: {AZUL_ESCURO}; text-decoration: underline; }}
QLabel#aviso {{ border-radius: 8px; padding: 10px 12px; font-size: 13px; }}
QLabel#aviso[tipo="erro"] {{ background: #FBEDED; color: {tema.ERRO}; }}
QLabel#aviso[tipo="ok"] {{ background: #E9F4EE; color: {tema.OK}; }}
QLabel#aviso[tipo="info"] {{ background: #EEF3F8; color: {AZUL_ESCURO}; }}
QLabel#conexao {{ color: {tema.TINTA3}; font-size: 12px; }}
"""


class _Tarefa(QObject):
    """Roda uma função do KeyAuth fora da thread da tela."""
    pronto = Signal(object, object)          # (resultado, erro)

    def __init__(self, funcao):
        super().__init__()
        self.funcao = funcao

    def rodar(self):
        try:
            self.pronto.emit(self.funcao(), None)
        except auth.ErroLogin as e:
            self.pronto.emit(None, str(e))
        except Exception as e:                 # nunca derruba a tela
            self.pronto.emit(None, f"Erro inesperado no login: {e}")


class _Ponte(QObject):
    """Vive na thread da tela: recebe o resultado e chama a função de volta
    ali (mexer na tela fora dessa thread derruba o Qt)."""
    def __init__(self, funcao, parent):
        super().__init__(parent)
        self.funcao = funcao

    @Slot(object, object)
    def receber(self, resultado, erro):
        self.funcao(resultado, erro)


class TelaLogin(QDialog):
    def __init__(self, keyauth=None, parent=None):
        super().__init__(parent)
        self.setObjectName("telaLogin")
        self.setWindowTitle("Automação - Fluxos RH  |  Entrar")
        self.setStyleSheet(ESTILO)
        self.setFixedSize(980, 620)
        self.ka = keyauth or auth.KeyAuth()
        self.usuario = None
        self._threads = []
        self.cfg = QSettings("PrefeituraNovaLima", "FluxosRH")

        raiz = QHBoxLayout(self)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.setSpacing(0)
        raiz.addWidget(self._painel_marca())
        raiz.addWidget(self._painel_formularios(), 1)

        usuario_salvo = self.cfg.value("login/usuario", "", type=str)
        if usuario_salvo:
            self.e_usuario.setText(usuario_salvo)
            self.lembrar.setChecked(True)
        self.trocar(0)
        self._conectar()

    # ------------------------------------------------------------ esquerda
    def _painel_marca(self):
        f = QFrame()
        f.setObjectName("marca")
        f.setFixedWidth(400)
        v = QVBoxLayout(f)
        v.setContentsMargins(44, 48, 44, 36)
        v.setSpacing(0)
        orgao = QLabel("Prefeitura de Nova Lima")
        orgao.setObjectName("marcaOrgao")
        v.addWidget(orgao)
        v.addSpacing(10)
        nome = QLabel("Fluxos RH")
        nome.setObjectName("marcaNome")
        v.addWidget(nome)
        v.addSpacing(6)
        sub = QLabel("Automação da Divisão de Benefícios")
        sub.setObjectName("marcaSub")
        v.addWidget(sub)
        v.addSpacing(40)
        linha = QFrame()
        linha.setObjectName("linhaMarca")
        v.addWidget(linha)
        v.addSpacing(28)
        for svg, titulo, texto in (
                (CARTAO, "Benefícios VR/VA", "Carga mensal do vale-refeição\ne do vale-alimentação"),
                (LUPA, "Consulta de Admissões", "Novos admitidos do SIGA e\nfichas da e-Turmalina")):
            h = QHBoxLayout()
            h.setSpacing(16)
            ic = QLabel()
            ic.setPixmap(_icone(svg, 44).pixmap(44, 44))
            ic.setFixedSize(48, 48)
            ic.setAlignment(Qt.AlignmentFlag.AlignTop)
            h.addWidget(ic, 0, Qt.AlignmentFlag.AlignTop)
            col = QVBoxLayout()
            col.setSpacing(3)
            t = QLabel(titulo)
            t.setObjectName("marcaItem")
            s = QLabel(texto)
            s.setObjectName("marcaItemSub")
            col.addWidget(t)
            col.addWidget(s)
            h.addLayout(col, 1)
            v.addLayout(h)
            v.addSpacing(22)
        v.addStretch(1)
        rod = QLabel(f"Versão {auth.VERSAO}. Acesso só para servidores autorizados.")
        rod.setObjectName("marcaRodape")
        rod.setWordWrap(True)
        v.addWidget(rod)
        return f

    # ------------------------------------------------------------ direita
    def _painel_formularios(self):
        w = QWidget()
        w.setObjectName("pagina")
        v = QVBoxLayout(w)
        v.setContentsMargins(64, 44, 64, 28)
        v.setSpacing(0)

        sel = QFrame()
        sel.setObjectName("seletor")
        sel.setFixedWidth(300)
        hs = QHBoxLayout(sel)
        hs.setContentsMargins(4, 4, 4, 4)
        hs.setSpacing(4)
        self.grupo = QButtonGroup(self)
        for i, t in enumerate(("Entrar", "Criar conta")):
            b = QPushButton(t)
            b.setObjectName("aba")
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            self.grupo.addButton(b, i)
            hs.addWidget(b)
        self.grupo.idClicked.connect(self.trocar)
        v.addWidget(sel)
        v.addSpacing(30)

        self.pilha = QStackedWidget()
        self.pilha.setObjectName("pilhaLogin")
        self.pilha.addWidget(self._form_entrar())
        self.pilha.addWidget(self._form_cadastro())
        v.addWidget(self.pilha, 1)

        self.conexao = QLabel("Conectando ao servidor de login...")
        self.conexao.setObjectName("conexao")
        v.addWidget(self.conexao)
        return w

    def _campo(self, dica="", senha=False):
        e = QLineEdit()
        e.setObjectName("campo")
        e.setPlaceholderText(dica)
        if senha:
            e.setEchoMode(QLineEdit.EchoMode.Password)
            acao = QAction(_icone(OLHO, 18), "Mostrar a senha", e)
            e.addAction(acao, QLineEdit.ActionPosition.TrailingPosition)

            def alternar():
                visivel = e.echoMode() == QLineEdit.EchoMode.Password
                e.setEchoMode(QLineEdit.EchoMode.Normal if visivel else QLineEdit.EchoMode.Password)
                acao.setIcon(_icone(OLHO_FECHADO if visivel else OLHO, 18))
                acao.setText("Esconder a senha" if visivel else "Mostrar a senha")
            acao.triggered.connect(alternar)
        return e

    def _rotulo(self, texto):
        r = QLabel(texto)
        r.setObjectName("rotulo")
        return r

    def _aviso(self):
        a = QLabel()
        a.setObjectName("aviso")
        a.setWordWrap(True)
        a.hide()
        return a

    def _form_entrar(self):
        w = QWidget()
        w.setObjectName("pagina")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        t = QLabel("Entrar")
        t.setObjectName("tituloForm")
        v.addWidget(t)
        v.addSpacing(6)
        s = QLabel("Use o usuário e a senha da sua conta.")
        s.setObjectName("subForm")
        v.addWidget(s)
        v.addSpacing(26)
        v.addWidget(self._rotulo("Usuário"))
        v.addSpacing(6)
        self.e_usuario = self._campo("seu usuário")
        v.addWidget(self.e_usuario)
        v.addSpacing(16)
        v.addWidget(self._rotulo("Senha"))
        v.addSpacing(6)
        self.e_senha = self._campo("sua senha", senha=True)
        v.addWidget(self.e_senha)
        # código da verificação em duas etapas: só aparece se a conta pedir
        self.box_2fa = QWidget()
        self.box_2fa.setObjectName("pagina")
        b2 = QVBoxLayout(self.box_2fa)
        b2.setContentsMargins(0, 16, 0, 0)
        b2.setSpacing(6)
        b2.addWidget(self._rotulo("Código de verificação (6 dígitos)"))
        self.e_2fa = self._campo("000000")
        self.e_2fa.setMaxLength(6)
        b2.addWidget(self.e_2fa)
        self.box_2fa.hide()
        v.addWidget(self.box_2fa)
        v.addSpacing(14)
        self.lembrar = QCheckBox("Lembrar meu usuário neste computador")
        self.lembrar.setObjectName("lembrar")
        v.addWidget(self.lembrar)
        v.addSpacing(20)
        self.b_entrar = QPushButton("Entrar")
        self.b_entrar.setObjectName("acao")
        self.b_entrar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.b_entrar.clicked.connect(self.entrar)
        v.addWidget(self.b_entrar)
        v.addSpacing(12)
        self.aviso_entrar = self._aviso()
        v.addWidget(self.aviso_entrar)
        v.addStretch(1)
        h = QHBoxLayout()
        h.setSpacing(6)
        q = QLabel("Ainda não tem conta?")
        q.setObjectName("dica")
        h.addWidget(q)
        lk = QPushButton("Criar conta com a sua chave de acesso")
        lk.setObjectName("link")
        lk.setCursor(Qt.CursorShape.PointingHandCursor)
        lk.clicked.connect(lambda: self.trocar(1))
        h.addWidget(lk)
        h.addStretch(1)
        v.addLayout(h)
        for e in (self.e_usuario, self.e_senha, self.e_2fa):
            e.returnPressed.connect(self.entrar)
        return w

    def _form_cadastro(self):
        w = QWidget()
        w.setObjectName("pagina")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        t = QLabel("Criar conta")
        t.setObjectName("tituloForm")
        v.addWidget(t)
        v.addSpacing(6)
        s = QLabel("Você precisa da chave de acesso que o administrador enviou.")
        s.setObjectName("subForm")
        s.setWordWrap(True)
        v.addWidget(s)
        v.addSpacing(20)
        v.addWidget(self._rotulo("Chave de acesso"))
        v.addSpacing(6)
        self.c_chave = self._campo("cole aqui a chave recebida")
        v.addWidget(self.c_chave)
        v.addSpacing(14)
        v.addWidget(self._rotulo("Usuário"))
        v.addSpacing(6)
        self.c_usuario = self._campo("escolha um nome de usuário")
        v.addWidget(self.c_usuario)
        v.addSpacing(14)
        dupla = QHBoxLayout()
        dupla.setSpacing(12)
        for attr, rot, dica in (("c_senha", "Senha", "no mínimo 6 caracteres"),
                                ("c_senha2", "Repita a senha", "a mesma senha")):
            col = QVBoxLayout()
            col.setSpacing(6)
            col.addWidget(self._rotulo(rot))
            e = self._campo(dica, senha=True)
            setattr(self, attr, e)
            col.addWidget(e)
            dupla.addLayout(col)
        v.addLayout(dupla)
        v.addSpacing(22)
        self.b_cadastrar = QPushButton("Criar conta e entrar")
        self.b_cadastrar.setObjectName("acao")
        self.b_cadastrar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.b_cadastrar.clicked.connect(self.cadastrar)
        v.addWidget(self.b_cadastrar)
        v.addSpacing(12)
        self.aviso_cadastro = self._aviso()
        v.addWidget(self.aviso_cadastro)
        v.addStretch(1)
        h = QHBoxLayout()
        h.setSpacing(6)
        q = QLabel("Já tem conta?")
        q.setObjectName("dica")
        h.addWidget(q)
        lk = QPushButton("Entrar")
        lk.setObjectName("link")
        lk.setCursor(Qt.CursorShape.PointingHandCursor)
        lk.clicked.connect(lambda: self.trocar(0))
        h.addWidget(lk)
        h.addStretch(1)
        v.addLayout(h)
        for e in (self.c_chave, self.c_usuario, self.c_senha, self.c_senha2):
            e.returnPressed.connect(self.cadastrar)
        return w

    # ------------------------------------------------------------ ações
    def trocar(self, i):
        self.grupo.button(i).setChecked(True)
        self.pilha.setCurrentIndex(i)
        alvo = (self.e_senha if self.e_usuario.text() else self.e_usuario) if i == 0 else self.c_chave
        alvo.setFocus()

    def _mostrar(self, aviso, texto, tipo="erro"):
        aviso.setProperty("tipo", tipo)
        aviso.setText(texto)
        aviso.style().unpolish(aviso)
        aviso.style().polish(aviso)
        aviso.setVisible(bool(texto))

    def _em_segundo_plano(self, funcao, ao_terminar):
        th = QThread(self)
        tarefa = _Tarefa(funcao)
        tarefa.moveToThread(th)
        ponte = _Ponte(ao_terminar, self)
        th.started.connect(tarefa.rodar)
        tarefa.pronto.connect(ponte.receber, Qt.ConnectionType.QueuedConnection)
        tarefa.pronto.connect(th.quit)
        item = (th, tarefa, ponte)
        th.finished.connect(lambda: self._threads.remove(item) if item in self._threads else None)
        self._threads.append(item)
        th.start()

    def _ocupado(self, sim, botao=None, texto=""):
        for b in (self.b_entrar, self.b_cadastrar):
            b.setEnabled(not sim)
        for g in self.grupo.buttons():
            g.setEnabled(not sim)
        if botao is not None:
            botao.setText(texto)

    def _conectar(self):
        def fim(_, erro):
            if erro:
                self.conexao.setText("Sem conexão com o servidor de login. Ele tenta de novo ao entrar.")
                self.conexao.setStyleSheet(f"color: {tema.ERRO};")
            else:
                self.conexao.setText("Conectado ao servidor de login.")
                self.conexao.setStyleSheet(f"color: {tema.OK};")
        self._em_segundo_plano(self.ka.iniciar, fim)

    def entrar(self):
        if not self.b_entrar.isEnabled():
            return
        u, s = self.e_usuario.text().strip(), self.e_senha.text()
        codigo = self.e_2fa.text().strip() if self.box_2fa.isVisible() else None
        if not u or not s:
            self._mostrar(self.aviso_entrar, "Preencha o usuário e a senha.")
            (self.e_usuario if not u else self.e_senha).setFocus()
            return
        self._mostrar(self.aviso_entrar, "")
        self._ocupado(True, self.b_entrar, "Entrando...")

        def fim(usuario, erro):
            self._ocupado(False, self.b_entrar, "Entrar")
            if erro:
                if "duas etapas" in erro:
                    self.box_2fa.show()
                    self.e_2fa.setFocus()
                    self._mostrar(self.aviso_entrar, erro, "info")
                else:
                    self._mostrar(self.aviso_entrar, erro)
                    self.e_senha.selectAll()
                    self.e_senha.setFocus()
                return
            self.cfg.setValue("login/usuario", u if self.lembrar.isChecked() else "")
            self._concluir(usuario)
        self._em_segundo_plano(lambda: self.ka.entrar(u, s, codigo or None), fim)

    def cadastrar(self):
        if not self.b_cadastrar.isEnabled():
            return
        chave = self.c_chave.text().strip()
        u = self.c_usuario.text().strip()
        s, s2 = self.c_senha.text(), self.c_senha2.text()
        faltando = [(n, e) for n, e in (("a chave de acesso", self.c_chave), ("o usuário", self.c_usuario),
                                        ("a senha", self.c_senha)) if not e.text().strip()]
        if faltando:
            self._mostrar(self.aviso_cadastro, f"Preencha {faltando[0][0]}.")
            faltando[0][1].setFocus()
            return
        if len(u) < 3:
            self._mostrar(self.aviso_cadastro, "O usuário precisa ter pelo menos 3 caracteres.")
            self.c_usuario.setFocus()
            return
        if len(s) < 6:
            self._mostrar(self.aviso_cadastro, "A senha precisa ter pelo menos 6 caracteres.")
            self.c_senha.setFocus()
            return
        if s != s2:
            self._mostrar(self.aviso_cadastro, "As duas senhas não são iguais.")
            self.c_senha2.selectAll()
            self.c_senha2.setFocus()
            return
        self._mostrar(self.aviso_cadastro, "")
        self._ocupado(True, self.b_cadastrar, "Criando a conta...")

        def fim(usuario, erro):
            self._ocupado(False, self.b_cadastrar, "Criar conta e entrar")
            if erro:
                self._mostrar(self.aviso_cadastro, erro)
                return
            self.cfg.setValue("login/usuario", u)
            self._concluir(usuario)
        self._em_segundo_plano(lambda: self.ka.cadastrar(u, s, chave), fim)

    def _concluir(self, usuario):
        self.usuario = usuario
        self.accept()

    def closeEvent(self, ev):
        for th, *_ in list(self._threads):
            th.quit()
            th.wait(2000)
        super().closeEvent(ev)


def pedir_login(parent=None, keyauth=None):
    """Mostra a tela. Devolve (KeyAuth, Usuario) ou (None, None) se fechar."""
    tela = TelaLogin(keyauth, parent)
    if tela.exec() == QDialog.DialogCode.Accepted and tela.usuario:
        return tela.ka, tela.usuario
    return None, None
