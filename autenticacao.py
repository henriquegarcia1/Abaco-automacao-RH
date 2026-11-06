"""Login do programa pelo KeyAuth (keyauth.cc), aplicação "abacorh".

Feito a partir do keyauth.py do exemplo oficial (pasta keyauth), com duas
mudanças para funcionar dentro de uma janela:
  - em vez de fechar o programa (os._exit) quando algo dá errado, levanta
    ErroLogin com a mensagem em português, e a tela mostra;
  - usa só a biblioteca padrão para a conexão e o pacote "cryptography" para
    conferir a assinatura das respostas (o exemplo usa requests, pywin32 e
    discord-interactions).
O protocolo é o mesmo: API 1.3, respostas assinadas (Ed25519) com a chave
pública do KeyAuth, e o HWID é o SID do usuário do Windows, igual ao exemplo.
"""
import hashlib
import json
import os
import platform
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

NOME_APP = "abacorh"
OWNER_ID = "dvtNDlx6N3"
VERSAO = "1.0"                 # tem que ser a mesma do painel do KeyAuth
URL_API = "https://keyauth.win/api/1.3/"
# chave pública com que o KeyAuth assina as respostas (a mesma do keyauth.py)
CHAVE_PUBLICA = "5586b4bc69c7a4b487e4563a4cd96afd39140f919bd31cea7d1c6a1e8439422b"
TOLERANCIA_S = 25              # resposta mais velha que isso é recusada


class ErroLogin(Exception):
    """Falha no login, no cadastro ou na conexão. A mensagem já vem pronta
    para mostrar na tela."""


# mensagens do KeyAuth -> português (procura pelo trecho, sem maiúsculas)
_TRADUCOES = [
    ("2fa", "Esta conta usa verificação em duas etapas. Digite o código de 6 dígitos."),
    ("invalid username", "Usuário não encontrado."),
    ("username not found", "Usuário não encontrado."),
    ("password does not match", "Senha incorreta."),
    ("invalid password", "Senha incorreta."),
    ("username already taken", "Este nome de usuário já está em uso. Escolha outro."),
    ("username already exists", "Este nome de usuário já está em uso. Escolha outro."),
    ("key already used", "Esta chave de acesso já foi usada."),
    ("key not found", "Chave de acesso não encontrada. Confira se digitou certo."),
    ("invalid license", "Chave de acesso inválida. Confira se digitou certo."),
    ("invalid key", "Chave de acesso inválida. Confira se digitou certo."),
    ("license key is banned", "Esta chave de acesso foi bloqueada pelo administrador."),
    ("key is banned", "Esta chave de acesso foi bloqueada pelo administrador."),
    ("hwid must be", "O identificador deste computador é curto demais para o KeyAuth. "
                     "Avise o administrador."),
    ("hwid", "Esta conta está ligada a outro computador. Peça ao administrador "
             "para liberar (resetar o HWID) no painel do KeyAuth."),
    ("expired", "O acesso desta conta venceu. Peça uma nova chave ao administrador."),
    ("no active subscription", "Esta conta não tem acesso ativo. Peça uma chave ao administrador."),
    ("banned", "Esta conta foi bloqueada pelo administrador."),
    ("blacklisted", "Este computador foi bloqueado pelo administrador."),
    ("paused", "O acesso ao programa está pausado pelo administrador."),
    ("disabled", "O acesso ao programa está desativado pelo administrador."),
    ("not validated", "A sessão expirou. Entre de novo."),
    ("session not found", "A sessão expirou. Tente de novo."),
    ("invalid session", "A sessão expirou. Tente de novo."),
]


def traduzir(msg):
    m = (msg or "").lower()
    for trecho, pt in _TRADUCOES:
        if trecho in m:
            return pt
    return msg or "O servidor de login recusou o pedido."


def hwid():
    """Identificador do computador. No Windows é o SID do usuário, como no
    exemplo oficial (assim o painel do KeyAuth mostra o mesmo valor)."""
    sistema = platform.system()
    try:
        if sistema == "Windows":
            sem_janela = 0x08000000                # CREATE_NO_WINDOW
            saida = subprocess.run(["whoami", "/user", "/fo", "csv", "/nh"],
                                   capture_output=True, text=True, timeout=10,
                                   creationflags=sem_janela).stdout
            sid = saida.strip().split(",")[-1].strip().strip('"')
            if sid.startswith("S-"):
                return sid
        elif sistema == "Linux":
            with open("/etc/machine-id") as f:
                return f.read().strip()
        elif sistema == "Darwin":
            saida = subprocess.run("ioreg -l | grep IOPlatformSerialNumber", shell=True,
                                   capture_output=True, text=True, timeout=10).stdout
            return saida.split("=", 1)[1].replace(" ", "").strip('"\n')
    except Exception:
        pass
    return hashlib.sha256(f"{platform.node()}-{uuid.getnode():012x}".encode()).hexdigest()


def _hash_do_programa():
    """md5 do .exe (ou do script), como o getchecksum() do exemplo. Só vale
    se a conferência de hash estiver ligada no painel."""
    try:
        alvo = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(sys.argv[0])
        with open(alvo, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return ""


def _conferir_assinatura(corpo, assinatura, carimbo):
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
        from cryptography.exceptions import InvalidSignature
    except ImportError:
        raise ErroLogin("Falta o componente de segurança do login (pacote "
                        "'cryptography'). Rode o INICIAR.bat de novo.")
    if not assinatura or not carimbo:
        raise ErroLogin("A resposta do servidor de login veio sem assinatura. "
                        "Pode ser um proxy da rede alterando a conexão.")
    try:
        enviado = int(carimbo)
    except ValueError:
        raise ErroLogin("A resposta do servidor de login veio com horário inválido.")
    if time.time() - enviado > TOLERANCIA_S:
        raise ErroLogin("A resposta do servidor de login chegou atrasada. Confira "
                        "se o relógio do computador está certo e tente de novo.")
    try:
        chave = Ed25519PublicKey.from_public_bytes(bytes.fromhex(CHAVE_PUBLICA))
        chave.verify(bytes.fromhex(assinatura), carimbo.encode() + corpo)
    except (InvalidSignature, ValueError):
        raise ErroLogin("A resposta do servidor de login não confere (assinatura "
                        "inválida). Tente de novo; se continuar, avise o administrador.")


class Usuario:
    def __init__(self, info):
        self.nome = info.get("username", "")
        self.ip = info.get("ip", "")
        self.hwid = info.get("hwid") or ""
        self.criado = info.get("createdate", "")
        self.ultimo_login = info.get("lastlogin", "")
        self.assinaturas = info.get("subscriptions") or []
        primeira = self.assinaturas[0] if self.assinaturas else {}
        self.plano = primeira.get("subscription", "")
        self.expira = primeira.get("expiry", "")

    def expira_em(self):
        """Data de vencimento do acesso como texto dd/mm/aaaa ('' se não houver)."""
        try:
            return time.strftime("%d/%m/%Y", time.localtime(int(self.expira)))
        except (TypeError, ValueError):
            return ""


class KeyAuth:
    """Cliente do KeyAuth. Uso:
        ka = KeyAuth(); ka.iniciar(); usuario = ka.entrar("ana", "senha")
    Cada método levanta ErroLogin com a mensagem pronta para a tela."""

    def __init__(self, nome=NOME_APP, ownerid=OWNER_ID, versao=VERSAO, timeout=15):
        if len(ownerid) != 10:
            raise ErroLogin("O ownerid do KeyAuth está errado (precisa ter 10 caracteres).")
        self.nome, self.ownerid, self.versao = nome, ownerid, versao
        self.timeout = timeout
        self.sessao = ""
        self.usuario = None
        self.link_atualizacao = ""

    # ---- conexão
    def _pedir(self, dados, assinado=True):
        dados = {**dados, "name": self.nome, "ownerid": self.ownerid}
        if self.sessao and "sessionid" not in dados and dados["type"] != "init":
            dados["sessionid"] = self.sessao
        corpo = urllib.parse.urlencode(dados).encode()
        pedido = urllib.request.Request(
            URL_API, data=corpo, method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": "FluxosRH/" + self.versao})
        try:
            with urllib.request.urlopen(pedido, timeout=self.timeout) as r:
                bruto = r.read()
                assinatura = r.headers.get("x-signature-ed25519")
                carimbo = r.headers.get("x-signature-timestamp")
        except urllib.error.HTTPError as e:
            raise ErroLogin(f"O servidor de login respondeu com erro {e.code}. Tente de novo.")
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            raise ErroLogin("Sem conexão com o servidor de login (keyauth.win). Confira "
                            f"a internet e tente de novo.\n({getattr(e, 'reason', e)})")
        texto = bruto.decode("utf-8", "replace")
        if texto.strip() == "KeyAuth_Invalid":
            raise ErroLogin("A aplicação 'abacorh' não existe no KeyAuth (nome ou ownerid errado).")
        if assinado:
            _conferir_assinatura(bruto, assinatura, carimbo)
        try:
            return json.loads(texto)
        except ValueError:
            raise ErroLogin("Resposta inesperada do servidor de login.")

    def _exigir_sessao(self):
        if not self.sessao:
            self.iniciar()

    # ---- operações
    def iniciar(self):
        r = self._pedir({"type": "init", "ver": self.versao, "hash": _hash_do_programa()})
        if r.get("message") == "invalidver":
            self.link_atualizacao = r.get("download") or ""
            raise ErroLogin("Há uma versão nova do programa. "
                            + ("Baixe pelo link: " + self.link_atualizacao
                               if self.link_atualizacao else
                               "Peça a versão nova ao administrador."))
        if not r.get("success"):
            raise ErroLogin(traduzir(r.get("message")))
        self.sessao = r.get("sessionid", "")
        return True

    def entrar(self, usuario, senha, codigo_2fa=None):
        self._exigir_sessao()
        dados = {"type": "login", "username": usuario, "pass": senha, "hwid": hwid()}
        if codigo_2fa:
            dados["code"] = codigo_2fa
        r = self._pedir(dados)
        if not r.get("success"):
            raise ErroLogin(traduzir(r.get("message")))
        self.usuario = Usuario(r.get("info") or {})
        return self.usuario

    def cadastrar(self, usuario, senha, chave):
        self._exigir_sessao()
        r = self._pedir({"type": "register", "username": usuario, "pass": senha,
                         "key": chave, "hwid": hwid()})
        if not r.get("success"):
            raise ErroLogin(traduzir(r.get("message")))
        self.usuario = Usuario(r.get("info") or {})
        return self.usuario

    def entrar_com_chave(self, chave, codigo_2fa=None):
        self._exigir_sessao()
        dados = {"type": "license", "key": chave, "hwid": hwid()}
        if codigo_2fa:
            dados["code"] = codigo_2fa
        r = self._pedir(dados)
        if not r.get("success"):
            raise ErroLogin(traduzir(r.get("message")))
        self.usuario = Usuario(r.get("info") or {})
        return self.usuario

    def renovar(self, usuario, chave):
        """Usa uma chave nova para estender o acesso de uma conta (upgrade)."""
        self._exigir_sessao()
        r = self._pedir({"type": "upgrade", "username": usuario, "key": chave})
        if not r.get("success"):
            raise ErroLogin(traduzir(r.get("message")))
        return r.get("message", "")

    def sessao_valida(self):
        try:
            return bool(self.sessao) and bool(self._pedir({"type": "check"}).get("success"))
        except ErroLogin:
            return False

    def registrar(self, mensagem):
        """Manda uma linha para o registro (logs) do painel do KeyAuth."""
        try:
            self._pedir({"type": "log", "pcuser": os.getenv("USERNAME") or os.getenv("USER") or "",
                         "message": mensagem}, assinado=False)
        except ErroLogin:
            pass

    def sair(self):
        try:
            if self.sessao:
                self._pedir({"type": "logout"})
        except ErroLogin:
            pass
        self.sessao, self.usuario = "", None
