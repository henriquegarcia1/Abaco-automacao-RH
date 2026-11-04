# -*- coding: utf-8 -*-
"""
==========================================================================
  NÚCLEO DA AUTOMAÇÃO ÁBACO + EXCEL  —  tudo o que o programa FAZ
==========================================================================

Não tem tela: a interface Qt6 (abaco_automacao.py) chama estas funções
passando um objeto `ui` com:

    ui.log(texto)                 mostra uma linha no registro
    ui.perguntar(texto, opcoes)   espera a pessoa escolher; devolve a opção
    ui.escolher_arquivo(texto)    caminho escolhido, ou "" se desistiu
    ui.cancelado                  True quando a pessoa clicou em Cancelar

As funções que abrem o navegador demoram e rodam fora da thread da tela.
Quando a pessoa cancela, `ui.perguntar` levanta Cancelado.

--------------------------------------------------------------------------
TRANSPARÊNCIA (vale para TODAS as funções):
  - Nada roda escondido. O navegador fica visível; na observação há um
    aviso "🔴 Observando" fixo na tela.
  - Só acompanha o que acontece DENTRO do navegador que o programa abre.
  - Na observação, registra o PADRÃO (de qual campo se copiou, quanto
    tempo levou), NÃO o conteúdo copiado — nada de CPF, valores, etc.
  - A senha da Ábaco é digitada pela pessoa na TELA OFICIAL da Ábaco e
    nunca passa por este programa nem é guardada.
==========================================================================
"""

import calendar
import copy
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

PASTA = Path(__file__).resolve().parent
CONFIG = PASTA / "config.json"

PADRAO = {
    "url_abaco": "https://www.abaco.com.br",   # troque pela URL real
    "arquivo_fluxo": "fluxo_gravado.py",
    "planilha": "dados.xlsx",
    "relatorio": "relatorio_observacao",       # gera .json e .md
    "mostrar_navegador": True,
    "pasta_entrada": "entrada",                # relatórios do mês
    "pasta_saida": "saida",                    # planilhas geradas
    # Conferir: seletores da tela de consulta do servidor na Ábaco.
    # Saem do fluxo gravado (ex.: "#txtMatricula").
    "tela_consulta": {
        "campo_matricula": "",
        "botao_pesquisar": "",
        "espera_ms": 1500,
        "campos": {"nome": "", "data_admissao": "", "cargo": "",
                   "jornada": "", "secretaria": ""},
    },
    # regras de VR/VA: preenchidas na 1ª vez a partir de regras_beneficios.py
    "beneficios": {},
}

# bibliotecas que o programa usa: (módulo, pacote do pip)
BIBLIOTECAS = [("PySide6", "PySide6-Essentials"), ("playwright", "playwright"),
               ("pandas", "pandas"), ("openpyxl", "openpyxl"),
               ("xlrd", "xlrd"), ("lxml", "lxml")]

# relatórios do mês, na ordem em que são coletados
ARQUIVOS_ENTRADA = [
    ("carga_anterior", "Divisão", "Planilha de carga do mês anterior (a base: "
                                  "aba CARGA com o TIPO de cada servidor)"),
    ("servidores", "SIGA", "Relatório 'situação de afastamento' de todos os "
                           "servidores (situação, regime, secretaria, nomeação, "
                           "exoneração, cargos, salário, afastamento)"),
    ("ferias", "Ábaco", "Período de férias (janela de 3 meses)"),
    ("faltas", "Ábaco", "Faltas injustificadas do mês anterior à folha"),
    ("plantoes", "e-mail", "Plantões enviados pela secretaria (Saúde)"),
    ("conferencia_abaco", "robô", "Conferência dos novos admitidos na Ábaco"),
    ("ajustes", "Divisão", "Ajustes manuais: tipo, quantidade, desconto, "
                           "reembolso (opcional)"),
]
OPCIONAIS = {"conferencia_abaco": "opcional (passo 2)",
             "ajustes": "opcional (para resolver pendências)"}


class Cancelado(Exception):
    """A pessoa cancelou a tarefa."""


# ==========================================================================
# Configuração
# ==========================================================================
def mesclar(base, extra):
    """Junta dicionários em profundidade (o config.json pode ser parcial)."""
    for k, v in extra.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            mesclar(base[k], v)
        else:
            base[k] = v
    return base


def carregar_config():
    cfg = copy.deepcopy(PADRAO)
    if CONFIG.exists():
        try:
            mesclar(cfg, json.loads(CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass    # config com erro: segue com o padrão (a tela avisa)
    return cfg


def erro_no_config():
    """Mensagem de erro do config.json, ou "" se estiver ok."""
    if not CONFIG.exists():
        return ""
    try:
        json.loads(CONFIG.read_text(encoding="utf-8"))
        return ""
    except Exception as e:
        return str(e)


def salvar_config(cfg):
    CONFIG.write_text(json.dumps(cfg, indent=2, ensure_ascii=False),
                      encoding="utf-8")


def regras_do_config(cfg):
    """Regras de benefícios completas; grava no config.json para edição."""
    import regras_beneficios as rb
    regras = rb.regras_completas(cfg.get("beneficios"))
    if cfg.get("beneficios") != regras:
        cfg["beneficios"] = regras
        salvar_config(cfg)
    return regras


def pasta_entrada(cfg):
    return PASTA / cfg["pasta_entrada"]


def pasta_saida(cfg):
    return PASTA / cfg["pasta_saida"]


def competencia_padrao():
    """O mês seguinte: o benefício é pago antecipado (regra C3)."""
    hoje = date.today()
    return (hoje.year + 1, 1) if hoje.month == 12 else (hoje.year, hoje.month + 1)


def mes_antes(ano, mes, n=1):
    total = ano * 12 + (mes - 1) - n
    return total // 12, total % 12 + 1


def seletores_configurados(cfg):
    tela = cfg.get("tela_consulta", {})
    return bool(tela.get("campo_matricula")) and \
        any(tela.get("campos", {}).values())


def _na_entrada(cfg, nome):
    pasta = pasta_entrada(cfg)
    achados = [pasta / f"{nome}{ext}" for ext in (".xlsx", ".xls", ".csv", ".xlsm")]
    achados = [p for p in achados if p.exists()]
    return max(achados, key=lambda p: p.stat().st_mtime) if achados else None


def arquivo_carga_anterior(cfg, ano, mes):
    """A planilha base: a que está na entrada ou, se não houver, a que o
    programa gerou na competência anterior (saida/AAAA-MM/'MM- MÊS - AAAA')."""
    arq = _na_entrada(cfg, "carga_anterior")
    if arq or not ano:
        return arq
    import regras_beneficios as rb
    a1, m1 = mes_antes(ano, mes, 1)
    gerada = pasta_saida(cfg) / f"{a1}-{m1:02d}" / rb.nome_planilha_carga(a1, m1)
    return gerada if gerada.exists() else None


def status_entrada(cfg, ano=None, mes=None):
    """[(nome, sistema, descrição, arquivo ou None, data de modificação)]"""
    itens = []
    for nome, sistema, descricao in ARQUIVOS_ENTRADA:
        if nome == "carga_anterior":
            arq = arquivo_carga_anterior(cfg, ano, mes)
        else:
            arq = _na_entrada(cfg, nome)
        quando = datetime.fromtimestamp(arq.stat().st_mtime) if arq else None
        itens.append((nome, sistema, descricao, arq, quando))
    return itens


def abrir_no_sistema(caminho):
    """Abre arquivo ou pasta no programa padrão do Windows."""
    caminho = Path(caminho)
    if hasattr(os, "startfile"):
        os.startfile(str(caminho))
    else:
        subprocess.Popen(["xdg-open", str(caminho)])


# ==========================================================================
# Instalação (o Python em si é instalado pelo INICIAR.bat)
# ==========================================================================
def _python_console():
    """python.exe ao lado do pythonw.exe (para os subprocessos terem saída)."""
    exe = Path(sys.executable)
    alt = exe.with_name("python.exe")
    return str(alt if exe.name.lower() == "pythonw.exe" and alt.exists() else exe)


def _rodar(comando, ui):
    """Roda um comando mostrando a saída linha a linha. Devolve o código."""
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(comando, stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True,
                            encoding="utf-8", errors="replace",
                            creationflags=flags, env=env, cwd=PASTA)
    for linha in proc.stdout:
        if linha.strip():
            ui.log(linha.rstrip())
        if ui.cancelado:
            proc.terminate()
            raise Cancelado()
    return proc.wait()


def bibliotecas_instaladas():
    """{pacote: True/False}"""
    return {pip: importlib.util.find_spec(mod) is not None
            for mod, pip in BIBLIOTECAS}


# navegadores do próprio Windows que o Playwright sabe controlar; usados
# quando o Chromium não pôde ser baixado (rede com proxy, bloqueio etc.)
NAVEGADORES_WINDOWS = [
    ("msedge", "Microsoft Edge", r"Microsoft\Edge\Application\msedge.exe"),
    ("chrome", "Google Chrome", r"Google\Chrome\Application\chrome.exe"),
]


def navegador_do_windows():
    """(canal, nome) do Edge/Chrome instalado no Windows, ou (None, None)."""
    bases = [os.environ.get(v) for v in ("PROGRAMFILES(X86)", "PROGRAMFILES",
                                         "LOCALAPPDATA")]
    for canal, nome, relativo in NAVEGADORES_WINDOWS:
        if any(b and Path(b, relativo).exists() for b in bases):
            return canal, nome
    return None, None


def chromium_baixado():
    """True se o Chromium do Playwright está baixado (demora ~1 s)."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            return Path(p.chromium.executable_path).exists()
    except Exception:
        return False


def situacao_navegador():
    """(ok, texto) sobre o navegador que o robô vai usar."""
    if chromium_baixado():
        return True, "Chromium instalado."
    _, nome = navegador_do_windows()
    if nome:
        return True, (f"Chromium não baixado; o robô usa o {nome} do Windows. "
                      f"Para baixar o Chromium, use a tela Instalação.")
    return False, "Não instalado. Abra a tela Instalação ou rode o INICIAR.bat."


def navegador_instalado():
    """True se há navegador para o robô (Chromium ou Edge/Chrome do Windows)."""
    return situacao_navegador()[0]


def instalar_dependencias(ui):
    py = _python_console()
    ui.log(">> Instalando bibliotecas Python...")
    codigo = _rodar([py, "-m", "pip", "install", "-r",
                     str(PASTA / "requirements.txt"), "--prefer-binary",
                     "--disable-pip-version-check", "--no-input"], ui)
    if codigo != 0:
        raise RuntimeError(f"A instalação das bibliotecas falhou (código {codigo}). "
                           f"Verifique a internet e tente de novo.")
    ui.log(">> Baixando o navegador controlado (Chromium)...")
    codigo = _rodar([py, "-m", "playwright", "install", "chromium"], ui)
    if codigo != 0:
        _, nome = navegador_do_windows()
        if not nome:
            raise RuntimeError(f"O download do navegador falhou (código {codigo}) e "
                               f"não há Edge nem Chrome no computador. "
                               f"Verifique a internet e tente de novo.")
        ui.log(f"Não consegui baixar o Chromium (código {codigo}). "
               f"Sem problema: o robô vai usar o {nome}.")
    ui.log("Instalação concluída.")


def abrir_navegador(p, cfg):
    """Abre o Chromium do Playwright; se ele faltar ou não abrir, usa o
    Edge/Chrome do Windows. Sem nenhum, explica em vez de quebrar."""
    headless = not cfg["mostrar_navegador"]
    try:
        return p.chromium.launch(headless=headless)
    except Exception as e:
        erro = e
    canal, _ = navegador_do_windows()
    if canal:
        try:
            return p.chromium.launch(channel=canal, headless=headless)
        except Exception:
            pass
    if "Executable doesn't exist" in str(erro):
        raise RuntimeError("O navegador controlado não está instalado. "
                           "Use a tela Instalação ou rode o INICIAR.bat.")
    raise RuntimeError(f"Não consegui abrir o navegador: "
                       f"{str(erro).splitlines()[0]}")


def _playwright():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError:
        raise RuntimeError("Bibliotecas não instaladas. Use a tela "
                           "Instalação ou rode o INICIAR.bat.")


def _fechar(navegador):
    try:
        navegador.close()
    except Exception:
        pass


# ==========================================================================
# OBSERVAR — documenta como a pessoa trabalha (etapa de diagnóstico)
# ==========================================================================

# Script injetado em cada página: escuta copiar/colar e mostra o painel
# de observação. Registra PADRÃO (de onde copiou, tamanho), não conteúdo.
JS_OBSERVACAO = r"""
(() => {
  if (window.__obsAtivo) return;
  window.__obsAtivo = true;

  let n = 0;
  function bump() {
    n++; const c = document.getElementById('obsCont');
    if (c) c.innerText = n + ' eventos';
  }
  function reg(ev) { try { window.obsRegistrar(ev); } catch (e) {} bump(); }

  function descrever(el) {
    if (!el) return 'desconhecido';
    try {
      const rot = el.getAttribute && (el.getAttribute('aria-label')
                  || el.getAttribute('name') || el.getAttribute('placeholder'));
      if (rot) return rot;
      if (el.id) {
        const lab = document.querySelector('label[for="' + el.id + '"]');
        if (lab && lab.innerText) return lab.innerText.trim().slice(0, 40);
        return el.id;
      }
    } catch (e) {}
    return (el.tagName || 'elemento').toLowerCase();
  }

  // ouvintes ligados JÁ no documento (sobrevivem a trocas de tela)
  document.addEventListener('copy', () => {
    let tam = 0, el = document.activeElement;
    const ehCampo = el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA'
                  || el.isContentEditable);
    try {
      if (ehCampo) {
        // copiou de um campo: o tamanho vem da seleção do próprio campo
        if (typeof el.selectionStart === 'number'
            && el.selectionEnd > el.selectionStart) {
          tam = el.selectionEnd - el.selectionStart;
        } else {
          tam = (el.value || '').length;
        }
      } else {
        // copiou de texto solto na página
        const s = window.getSelection();
        tam = (s && s.toString()) ? s.toString().length : 0;
        if (s && s.anchorNode) el = s.anchorNode.parentElement || el;
      }
    } catch (e) {}
    reg({ tipo: 'copiou', origem: descrever(el), tamanho: tam });
  }, true);

  document.addEventListener('paste', () => {
    reg({ tipo: 'colou', destino: descrever(document.activeElement) });
  }, true);

  // painel visível; (re)injetado quando o corpo existe e se a página
  // reconstruir o DOM (sistemas que re-renderizam a tela toda)
  function injetarPainel() {
    if (!document.body || document.getElementById('obsBox')) return;
    const box = document.createElement('div');
    box.id = 'obsBox';
    box.style.cssText = 'position:fixed;right:12px;bottom:12px;'
      + 'z-index:2147483647;background:#1e293b;color:#fff;padding:10px;'
      + 'border-radius:8px;font:13px sans-serif;'
      + 'box-shadow:0 2px 12px rgba(0,0,0,.45);width:230px';
    box.innerHTML =
      '<div style="margin-bottom:6px">🔴 Observando (pode anotar)</div>'
      + '<input id="obsNota" placeholder="o que está fazendo..." '
      + 'style="width:100%;box-sizing:border-box;padding:5px;border:0;'
      + 'border-radius:4px;margin-bottom:6px">'
      + '<button id="obsBtn" style="width:100%;padding:5px;border:0;'
      + 'border-radius:4px;background:#38bdf8;cursor:pointer">Anotar</button>'
      + '<div id="obsCont" style="margin-top:6px;opacity:.7">'
      + n + ' eventos</div>';
    document.body.appendChild(box);
    const btn = box.querySelector('#obsBtn');
    const inp = box.querySelector('#obsNota');
    btn.addEventListener('click', () => {
      const t = inp.value.trim();
      if (t) { reg({ tipo: 'nota', texto: t }); inp.value = ''; }
    });
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') btn.click(); });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injetarPainel);
  } else {
    injetarPainel();
  }
  // garante o painel mesmo em telas que se redesenham
  setInterval(injetarPainel, 1500);
})();
"""


def observar(cfg, ui):
    """Abre a Ábaco e registra o padrão de trabalho até o navegador fechar.
    Devolve o caminho do relatório .md."""
    sync_playwright = _playwright()
    eventos = []

    def registrar(ev):
        try:
            ev = dict(ev)
        except Exception:
            ev = {"tipo": "?", "bruto": str(ev)}
        ev["hora"] = datetime.now().strftime("%H:%M:%S")
        eventos.append(ev)
        resumo = ev.get("origem") or ev.get("destino") or ev.get("texto") \
            or ev.get("titulo") or ""
        ui.log(f"• {ev['hora']}  {ev.get('tipo', '?'):8} {resumo}")

    with sync_playwright() as p:
        navegador = abrir_navegador(p, cfg)
        context = navegador.new_context()
        context.expose_function("obsRegistrar", registrar)
        context.add_init_script(JS_OBSERVACAO)

        def on_nav(pg, frame):
            if frame != pg.main_frame:
                return
            try:
                titulo_pg = pg.title()
            except Exception:
                titulo_pg = ""
            registrar({"tipo": "navegou", "titulo": titulo_pg, "url": frame.url})

        def wire(pg):
            pg.on("framenavigated", lambda frame: on_nav(pg, frame))

        context.on("page", wire)
        page = context.new_page()
        wire(page)

        fechado = {"v": False}
        context.on("close", lambda: fechado.__setitem__("v", True))
        page.on("close", lambda: fechado.__setitem__("v", True))

        ui.log(f"Abrindo {cfg['url_abaco']}")
        page.goto(cfg["url_abaco"])
        ui.log("Observando... trabalhe normalmente. FECHE o navegador no fim.")

        try:
            while not fechado["v"] and not ui.cancelado:
                page.wait_for_timeout(500)
        except Exception:
            pass  # navegador fechado encerra o laço
        _fechar(navegador)

    return gerar_relatorio(cfg, eventos, ui)


def gerar_relatorio(cfg, eventos, ui):
    base = PASTA / cfg["relatorio"]
    (base.with_suffix(".json")).write_text(
        json.dumps(eventos, indent=2, ensure_ascii=False), encoding="utf-8")

    # ---- resumo ----
    total = len(eventos)
    copias = [e for e in eventos if e.get("tipo") == "copiou"]
    colagens = [e for e in eventos if e.get("tipo") == "colou"]
    navs = [e for e in eventos if e.get("tipo") == "navegou"]
    notas = [e for e in eventos if e.get("tipo") == "nota"]

    # telas distintas por título (ou url)
    telas = {}
    for e in navs:
        chave = e.get("titulo") or e.get("url") or "?"
        telas[chave] = telas.get(chave, 0) + 1

    # campos de onde mais se copiou (candidatos fortes a automação)
    origens = {}
    for e in copias:
        o = e.get("origem", "?")
        origens[o] = origens.get(o, 0) + 1

    dur = ""
    if eventos:
        dur = f"{eventos[0].get('hora','?')} → {eventos[-1].get('hora','?')}"

    linhas = []
    linhas.append("# Relatório de Observação — Ábaco\n")
    linhas.append(f"Gerado em: {datetime.now():%d/%m/%Y %H:%M}  ")
    linhas.append(f"Período observado: {dur}\n")
    linhas.append("## Resumo\n")
    linhas.append(f"- Eventos registrados: **{total}**")
    linhas.append(f"- Vezes que copiou dados: **{len(copias)}**")
    linhas.append(f"- Vezes que colou dados: **{len(colagens)}**")
    linhas.append(f"- Mudanças de tela: **{len(navs)}**")
    linhas.append(f"- Anotações da pessoa: **{len(notas)}**\n")

    if telas:
        linhas.append("## Telas visitadas (quantas vezes)\n")
        for nome, q in sorted(telas.items(), key=lambda x: -x[1]):
            linhas.append(f"- {nome} — {q}x")
        linhas.append("")

    if origens:
        linhas.append("## De onde mais se copiou (candidatos a automação)\n")
        for nome, q in sorted(origens.items(), key=lambda x: -x[1]):
            linhas.append(f"- campo \"{nome}\" — {q}x")
        linhas.append("")

    if notas:
        linhas.append("## Anotações da pessoa (passo a passo narrado)\n")
        for e in notas:
            linhas.append(f"- {e['hora']} — {e['texto']}")
        linhas.append("")

    # leitura automática simples para orientar a decisão
    linhas.append("## Leitura inicial\n")
    repetidos = [f'"{n}" ({q}x)' for n, q in origens.items() if q >= 3]
    if repetidos:
        linhas.append("Há cópias repetidas dos mesmos campos: "
                      + ", ".join(repetidos)
                      + ". Repetição é o principal sinal de que automatizar "
                      + "vale a pena.")
    elif copias:
        linhas.append("Houve cópias, mas pouca repetição do mesmo campo. "
                      "Vale observar mais algumas vezes antes de decidir.")
    else:
        linhas.append("Não houve cópias registradas nesta sessão. Se o "
                      "trabalho é mais de digitação/consulta, me conte que "
                      "ajusto o que observar.")
    linhas.append("")
    linhas.append("## Linha do tempo (eventos)\n")
    for e in eventos:
        resumo = e.get("origem") or e.get("destino") or e.get("texto") \
            or e.get("titulo") or e.get("url") or ""
        linhas.append(f"- {e.get('hora','?')}  `{e.get('tipo','?')}`  {resumo}")

    md = base.with_suffix(".md")
    md.write_text("\n".join(linhas), encoding="utf-8")
    ui.log(f"Relatório gerado: {md}")
    ui.log(f"Resumo: {total} eventos | {len(copias)} cópias | "
           f"{len(navs)} telas | {len(notas)} anotações")
    return md


# ==========================================================================
# GRAVAR FLUXO  (depois de decidir automatizar)
# ==========================================================================
def gravar(cfg, ui):
    """Abre o gravador do Playwright; devolve o arquivo com o fluxo."""
    _playwright()
    saida = PASTA / cfg["arquivo_fluxo"]
    ui.log(f"Gravando em {saida}. Faça login, execute a tarefa UMA vez e "
           f"FECHE o navegador ao terminar.")
    comando = [_python_console(), "-m", "playwright", "codegen",
               "--target", "python", "--output", str(saida)]
    canal, nome = navegador_do_windows()
    if canal and not chromium_baixado():
        ui.log(f"Chromium não baixado; gravando com o {nome}.")
        comando += ["--channel", canal]
    codigo = _rodar(comando + [cfg["url_abaco"]], ui)
    if codigo != 0:
        raise RuntimeError(f"A gravação terminou com problema (código {codigo}).")
    ui.log(f"Fluxo salvo em: {saida}")
    return saida


# ==========================================================================
# EXTRAIR TABELA DA TELA PARA EXCEL
# ==========================================================================
def extrair(cfg, ui):
    """Lê tabelas da tela da Ábaco, página por página. Devolve o .xlsx."""
    sync_playwright = _playwright()
    import pandas as pd

    destino = PASTA / cfg["planilha"]
    todas = []
    with sync_playwright() as p:
        navegador = abrir_navegador(p, cfg)
        page = navegador.new_page()
        ui.log(f"Abrindo {cfg['url_abaco']}")
        page.goto(cfg["url_abaco"])
        pagina = 1
        try:
            while True:
                r = ui.perguntar(
                    "Faça login e navegue até a tela com os dados (aplique "
                    f"filtros se precisar). Quando a página {pagina} estiver "
                    "na tela, clique em Ler.", [f"Ler página {pagina}", "Terminar"])
                if r == "Terminar":
                    break
                try:
                    linhas = ler_tabela_da_tela(page)
                except Exception as e:
                    ui.log(f"[ERRO ao ler a tela] {str(e).splitlines()[0]}")
                    continue
                if not linhas:
                    ui.log("Nenhuma tabela encontrada. Se os dados não são uma "
                           "tabela comum, mande um print da tela para ajustar o leitor.")
                    continue
                ui.log(f"Li {len(linhas)} linhas na página {pagina}.")
                todas.extend(linhas)
                pagina += 1
        finally:
            _fechar(navegador)

    if not todas:
        ui.log("Nada foi extraído.")
        return None
    df = pd.DataFrame(todas)
    df.to_excel(destino, index=False)
    ui.log(f"{len(df)} linhas salvas em: {destino}")
    ui.log(f"Colunas: {', '.join(map(str, df.columns))}")
    return destino


def ler_tabela_da_tela(page):
    """Lê a primeira tabela HTML visível e devolve lista de dicionários."""
    page.wait_for_selector("table", timeout=15000)
    tabela = page.query_selector("table")
    if tabela is None:
        return []
    cabecalhos = [th.inner_text().strip()
                  for th in tabela.query_selector_all("thead th")]
    if not cabecalhos:
        primeira = tabela.query_selector("tr")
        if primeira:
            cabecalhos = [c.inner_text().strip()
                          for c in primeira.query_selector_all("th, td")]
    corpo = tabela.query_selector_all("tbody tr") \
        or tabela.query_selector_all("tr")[1:]
    linhas = []
    for tr in corpo:
        celulas = [td.inner_text().strip()
                   for td in tr.query_selector_all("td")]
        if not celulas:
            continue
        if cabecalhos and len(cabecalhos) == len(celulas):
            linhas.append(dict(zip(cabecalhos, celulas)))
        else:
            linhas.append({f"Coluna {i+1}": v for i, v in enumerate(celulas)})
    return linhas


# ==========================================================================
# BENEFÍCIOS VR/VA
# ==========================================================================
def gerar_modelos():
    import regras_beneficios as rb
    return rb.gerar_modelos(PASTA / "modelos")


def copiar_para_entrada(origem, cfg, nome):
    entrada = pasta_entrada(cfg)
    entrada.mkdir(parents=True, exist_ok=True)
    origem = Path(origem)
    destino = entrada / f"{nome}{origem.suffix.lower()}"
    shutil.copy2(origem, destino)
    return destino


def bombear(context):
    """Deixa o Playwright processar eventos (ex.: downloads) acumulados."""
    for pg in reversed(context.pages):
        try:
            pg.wait_for_timeout(1500)
            return True
        except Exception:
            continue
    return False


def relatorios_da_competencia(ano, mes):
    """(arquivo, sistema, o que emitir) — regras R3, R5, N5 e E5.
    Competência = mês de uso; a folha é a do mês anterior (C3)."""
    a1, m1 = mes_antes(ano, mes, 1)
    a2, m2 = mes_antes(ano, mes, 2)
    ap, mp = mes_antes(ano, mes, -1)
    fim = lambda a, m: f"{calendar.monthrange(a, m)[1]:02d}/{m:02d}/{a}"
    return [
        ("servidores", "SIGA",
         "Relatório 'Situação de afastamento' de TODOS os servidores, com as "
         "colunas: MATRICULA, NOME, CPF, SITUACAO, REGIME, NOME_SECRETARIA, "
         "NOMEACAO, EXONERACAO, CARGO_EFETIVO, CARGO_COMISSIONADO, SALARIO, "
         "INICIO_AFASTAMENTO, TERMINO_AFASTAMENTO e TIPO_AFASTAMENTO."),
        ("ferias", "Ábaco",
         f"Relatório de período de férias de 01/{m1:02d}/{a1} a "
         f"{fim(ap, mp)} (janela de 3 meses)."),
        ("faltas", "Ábaco",
         f"Relatório de faltas de {m2:02d}/{a2} (o mês anterior à folha de "
         f"{m1:02d}/{a1})."),
    ]


def coletar_relatorios(cfg, ano, mes, ui):
    """Guia a emissão dos relatórios e guarda os downloads na pasta de
    entrada com o nome certo. Por último, pede a planilha de plantões."""
    sync_playwright = _playwright()
    entrada = pasta_entrada(cfg)
    entrada.mkdir(parents=True, exist_ok=True)
    salvos = []
    baixados = []

    # a base do cálculo: a planilha de carga do mês anterior
    base = arquivo_carga_anterior(cfg, ano, mes)
    if base:
        ui.log(f">> Planilha do mês anterior: {base.name}")
    else:
        origem = ui.escolher_arquivo("Planilha de carga do MÊS ANTERIOR (ex.: "
                                     "'11- NOVEMBRO - 2026.xlsx')")
        if origem:
            salvo = copiar_para_entrada(origem, cfg, "carga_anterior")
            salvos.append(salvo)
            ui.log(f">> Planilha do mês anterior: {salvo}")
        else:
            ui.log("   Sem a planilha do mês anterior o cálculo não roda. "
                   "Adicione depois em 'Adicionar arquivo manualmente'.")

    def ao_baixar(download):
        baixados.append(download)

    with sync_playwright() as p:
        navegador = abrir_navegador(p, cfg)
        context = navegador.new_context(accept_downloads=True)
        context.on("page", lambda pg: pg.on("download", ao_baixar))
        page = context.new_page()
        page.on("download", ao_baixar)   # repetido é inofensivo: vale o último
        ui.log(f"Abrindo {cfg['url_abaco']}")
        page.goto(cfg["url_abaco"])
        try:
            ui.perguntar("Faça login no navegador que abriu (a senha é digitada "
                         "no site e não passa por este programa). Depois clique "
                         "em Continuar.", ["Continuar"])
            for nome, sistema, descricao in relatorios_da_competencia(ano, mes):
                baixados.clear()
                ui.log(f">> {nome.upper()} ({sistema}): {descricao}")
                while True:
                    r = ui.perguntar(
                        f"{nome.upper()} ({sistema})\n{descricao}\n\nGere o "
                        f"relatório EM EXCEL (não PDF) e baixe. Depois clique "
                        f"em Verificar download.",
                        ["Verificar download", "Escolher arquivo...", "Pular"])
                    if r == "Pular":
                        ui.log("   Pulado.")
                        break
                    if r == "Escolher arquivo...":
                        origem = ui.escolher_arquivo(f"Arquivo de {nome}")
                        if origem:
                            salvo = copiar_para_entrada(origem, cfg, nome)
                            salvos.append(salvo)
                            ui.log(f"   OK: {salvo}")
                            break
                        continue
                    if not bombear(context):
                        raise RuntimeError("O navegador foi fechado. Comece a "
                                           "coleta de novo.")
                    if not baixados:
                        ui.log("   Nenhum download percebido ainda.")
                        continue
                    d = baixados[-1]
                    baixados.clear()
                    ext = (Path(d.suggested_filename).suffix or ".xlsx").lower()
                    if ext == ".pdf":
                        ui.log("   Veio em PDF. Gere o relatório em Excel e baixe de novo.")
                        continue
                    destino = entrada / f"{nome}{ext}"
                    d.save_as(destino)
                    salvos.append(destino)
                    ui.log(f"   OK: {destino}")
                    break
        finally:
            _fechar(navegador)

    # R6: a planilha de plantões chega por e-mail, não sai do sistema
    origem = ui.escolher_arquivo("Planilha de PLANTÕES enviada pela secretaria "
                                 "(cancele para pular)")
    if origem:
        salvo = copiar_para_entrada(origem, cfg, "plantoes")
        salvos.append(salvo)
        ui.log(f">> PLANTÕES: {salvo}")
    ui.log("Coleta concluída. Próximo passo: Conferir (opcional) e Calcular.")
    return salvos


# ---- CONFERIR NOVOS ADMITIDOS NA ÁBACO  (robô de consulta — regra N1)
JS_LER_CAMPO = """e => e.tagName === 'SELECT'
    ? (e.selectedOptions[0] ? e.selectedOptions[0].text : '')
    : (('value' in e && e.tagName !== 'BUTTON') ? e.value : e.innerText)"""


def conferir_na_abaco(cfg, ano, mes, ui):
    """Digita cada matrícula nova na tela de consulta e lê os dados.
    Grava entrada/conferencia_abaco.xlsx e devolve o caminho."""
    import pandas as pd
    import regras_beneficios as rb
    sync_playwright = _playwright()

    tela = cfg["tela_consulta"]
    campos = {k: v for k, v in tela.get("campos", {}).items() if v}
    if not seletores_configurados(cfg):
        raise RuntimeError(
            "O robô ainda não sabe onde ficam os campos na tela de consulta da "
            "Ábaco. Grave uma consulta (Diagnóstico > Gravar fluxo) e preencha "
            "os seletores em Configurações > Robô de consulta.")

    regras = regras_do_config(cfg)
    novos = rb.novos_admitidos(pasta_entrada(cfg), ano, mes, regras,
                               arquivo_carga_anterior(cfg, ano, mes))
    if not novos:
        ui.log("Nenhum novo admitido no relatório de servidores para conferir.")
        return None
    ui.log(f"{len(novos)} matrícula(s) para conferir.")

    linhas = []
    with sync_playwright() as p:
        navegador = abrir_navegador(p, cfg)
        page = navegador.new_page()
        page.goto(cfg["url_abaco"])
        try:
            ui.perguntar(f"{len(novos)} matrícula(s) para conferir. Faça login, "
                         f"abra a TELA DE CONSULTA do servidor e clique em "
                         f"Começar. Dá para acompanhar o robô no navegador.",
                         ["Começar"])
            for i, (mat, nome) in enumerate(novos, 1):
                if ui.cancelado:
                    raise Cancelado()
                linha = {"Matrícula": mat}
                try:
                    page.fill(tela["campo_matricula"], mat)
                    if tela.get("botao_pesquisar"):
                        page.click(tela["botao_pesquisar"])
                    else:
                        page.press(tela["campo_matricula"], "Enter")
                    try:
                        page.wait_for_load_state("networkidle", timeout=10000)
                    except Exception:
                        pass
                    page.wait_for_timeout(int(tela.get("espera_ms", 1500)))
                    for campo, seletor in campos.items():
                        el = page.query_selector(seletor)
                        valor = el.evaluate(JS_LER_CAMPO).strip() if el else ""
                        linha[rb.ROTULOS[campo]] = valor
                    erro = ""
                    # proteção: se a tela não atualizou, o nome lido é de outra pessoa
                    lido, esperado = rb.norm(linha.get("Nome")), rb.norm(nome)
                    if lido and esperado and lido not in esperado \
                            and esperado not in lido:
                        erro = f"nome na tela ('{linha['Nome']}') não confere"
                        linha = {"Matrícula": mat}
                except Exception as e:
                    erro = str(e).splitlines()[0][:200]
                    linha = {"Matrícula": mat}
                linha["Erro"] = erro
                linhas.append(linha)
                ui.log(f"{i:3}/{len(novos)}  {mat:>10}  "
                       f"{'ERRO: ' + erro if erro else 'ok'}")
        finally:
            _fechar(navegador)

    destino = pasta_entrada(cfg) / "conferencia_abaco.xlsx"
    pd.DataFrame(linhas).to_excel(destino, index=False)
    erros = sum(1 for l in linhas if l["Erro"])
    ui.log(f"Conferência salva em: {destino}")
    ui.log(f"{len(linhas) - erros} ok, {erros} com erro (essas ficam com os "
           f"dados do SIGA). O cálculo usa os dados da Ábaco e lista as divergências.")
    return destino


# ---- CALCULAR VR/VA DA COMPETÊNCIA
def info_dias_uteis(cfg, ano, mes):
    """(dias úteis, [(data, feriado)], [(data, ponto facultativo sugerido)])"""
    import regras_beneficios as rb
    regras = regras_do_config(cfg)
    cal = rb.Calendario.do_config(ano - 1, ano + 1, regras)
    sugestoes = [(d, desc) for d, desc in rb.pontos_facultativos_comuns(ano).items()
                 if d.month == mes and d not in cal.folgas and d.weekday() < 5]
    return cal.dias_uteis_mes(ano, mes), cal.folgas_no_mes(ano, mes), sugestoes


def calcular_beneficios(cfg, ano, mes, dias_uteis, ui):
    """Calcula e grava as planilhas. Devolve (Resultado, [arquivos])."""
    import regras_beneficios as rb
    regras = regras_do_config(cfg)
    ui.log(f"Calculando VR/VA de {mes:02d}/{ano} com {dias_uteis} dias úteis...")
    res = rb.calcular(pasta_entrada(cfg), ano, mes, regras, dias_uteis=dias_uteis,
                      carga_anterior=arquivo_carga_anterior(cfg, ano, mes))
    try:
        arquivos = rb.salvar(res, pasta_saida(cfg) / res.rotulo, regras)
    except PermissionError as e:
        raise RuntimeError(f"Não consegui gravar: a planilha está aberta no "
                           f"Excel? Feche e tente de novo.\n{e}")
    ok = (res.vr["Status"] == "OK").sum()
    trava = (res.pendencias["Trava a importação"] == "Sim").sum()
    total = res.vr.loc[res.vr["Status"] == "OK", "Valor total"].fillna(0).sum()
    reais = f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    ui.log(f"VR: {len(res.vr)} na planilha, {ok} prontos ({reais})")
    ui.log(f"VA: {int((res.va['Crédito'].fillna(0) > 0).sum())} créditos")
    ui.log(f"Pendências: {len(res.pendencias)} ({trava} travam a importação)")
    for a in res.avisos:
        ui.log(f"• {a}")
    for a in arquivos:
        ui.log(f"Gerado: {a}")
    return res, arquivos
