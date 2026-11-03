#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================================
  AUTOMAÇÃO ÁBACO + EXCEL  —  programa único e interativo
==========================================================================

A ordem certa de trabalho (o menu segue esta ordem):

   [1] Instalar    -> prepara o computador (uma vez só)
   [2] OBSERVAR    -> PRIMEIRO PASSO: a pessoa trabalha e o programa
                      documenta o processo, para a gente decidir SE e
                      O QUE vale automatizar. Gera um relatório.
   [3] Gravar fluxo-> só depois de decidir automatizar: registra os
                      passos exatos de uma tarefa escolhida.
   [4] Extrair     -> executa: lê os dados da tela da Ábaco -> Excel.

--------------------------------------------------------------------------
TRANSPARÊNCIA (vale para TODOS os modos):
  - Nada roda escondido. Quem opera abre o programa e vê o que acontece;
    na observação há um aviso "🔴 Observando" fixo na tela.
  - Só acompanha o que acontece DENTRO do navegador que ELE abre, na
    Ábaco. Não lê o resto do computador.
  - Na observação, registra o PADRÃO (de qual campo se copiou, quanto
    tempo levou), NÃO o conteúdo copiado — nada de CPF, valores, etc.
  - A senha da Ábaco é digitada por você na TELA OFICIAL da Ábaco. A
    senha nunca passa por este programa nem é guardada.
--------------------------------------------------------------------------

Como começar:
    python abaco_automacao.py
==========================================================================
"""

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# --------------------------------------------------------------------------
# CONFIGURAÇÃO
# --------------------------------------------------------------------------
PASTA = Path(__file__).parent
CONFIG = PASTA / "config.json"

PADRAO = {
    "url_abaco": "https://www.abaco.com.br",   # troque pela URL real
    "arquivo_fluxo": "fluxo_gravado.py",
    "planilha": "dados.xlsx",
    "relatorio": "relatorio_observacao",       # gera .json e .md
    "mostrar_navegador": True,
}


# ==========================================================================
# Utilidades
# ==========================================================================
def carregar_config():
    cfg = dict(PADRAO)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def salvar_config(cfg):
    CONFIG.write_text(json.dumps(cfg, indent=2, ensure_ascii=False),
                      encoding="utf-8")


def pausa():
    input("\n  (pressione ENTER para voltar ao menu)")


def titulo(txt):
    print("\n" + "=" * 70)
    print("  " + txt)
    print("=" * 70)


# ==========================================================================
# [1] INSTALAR
# ==========================================================================
def instalar():
    titulo("[1] INSTALAR — prepara o computador (faça uma vez)")
    print("  Vai instalar: playwright, pandas, openpyxl + o navegador.\n")
    passos = [
        ("Instalando bibliotecas Python",
         [sys.executable, "-m", "pip", "install",
          "playwright==1.63.0", "pandas", "openpyxl"]),
        ("Baixando o navegador controlado (Chromium)",
         [sys.executable, "-m", "playwright", "install", "chromium"]),
    ]
    for descricao, comando in passos:
        print(f"  >> {descricao}...")
        try:
            subprocess.run(comando, check=True)
            print("     OK\n")
        except subprocess.CalledProcessError as e:
            print(f"     FALHOU: {e}\n")
            pausa()
            return
    print("  INSTALAÇÃO CONCLUÍDA. Agora comece pela opção [2] OBSERVAR.")
    pausa()


# ==========================================================================
# [2] OBSERVAR  — documenta como a pessoa trabalha (etapa de diagnóstico)
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


def observar(cfg):
    titulo("[2] OBSERVAR — documenta o trabalho (antes de automatizar)")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  [ERRO] Bibliotecas não instaladas. Rode a opção [1] primeiro.")
        pausa()
        return

    print("  Objetivo: entender COMO a pessoa trabalha, para decidir se e o")
    print("  que vale automatizar. Registra o padrão (telas visitadas, quando")
    print("  copia, tempo gasto) — NÃO o conteúdo copiado.\n")
    print("  Vai abrir o navegador. Faça login e trabalhe NORMALMENTE.")
    print("  Use o painel '🔴 Observando' para anotar o que está fazendo.")
    print("  Quando terminar, FECHE o navegador — o relatório é gerado.\n")
    if input("  Começar a observar? (s/n) ").strip().lower() != "s":
        return

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
        print(f"   • {ev['hora']}  {ev.get('tipo','?'):8} {resumo}")

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not cfg["mostrar_navegador"])
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

        print(f"\n  Abrindo {cfg['url_abaco']}")
        page.goto(cfg["url_abaco"])
        print("\n  Observando... trabalhe normalmente. FECHE o navegador no fim.\n")

        try:
            while not fechado["v"]:
                page.wait_for_timeout(500)
        except Exception:
            pass  # navegador fechado encerra o laço

        try:
            navegador.close()
        except Exception:
            pass

    gerar_relatorio(cfg, eventos)
    pausa()


def gerar_relatorio(cfg, eventos):
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
    linhas.append(f"# Relatório de Observação — Ábaco\n")
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

    (base.with_suffix(".md")).write_text("\n".join(linhas), encoding="utf-8")

    print("\n  RELATÓRIO GERADO:")
    print(f"    - {base.with_suffix('.md')}   (leitura fácil)")
    print(f"    - {base.with_suffix('.json')} (dados brutos)")
    print(f"\n  Resumo rápido: {total} eventos | {len(copias)} cópias | "
          f"{len(navs)} telas | {len(notas)} anotações")
    print("  Me mande o relatório .md que a gente decide juntos o que "
          "vale automatizar.")


# ==========================================================================
# [3] GRAVAR FLUXO  (depois de decidir automatizar)
# ==========================================================================
def gravar(cfg):
    titulo("[3] GRAVAR FLUXO — registra os passos de UMA tarefa escolhida")
    print("  Use isto DEPOIS de observar e decidir o que vale automatizar.\n")
    saida = PASTA / cfg["arquivo_fluxo"]
    print(f"  Site: {cfg['url_abaco']}")
    print(f"  Código gravado será salvo em: {saida}\n")
    print("  Faça login, execute a tarefa UMA vez do início ao fim,")
    print("  e FECHE o navegador ao terminar.\n")
    if input("  Começar a gravar? (s/n) ").strip().lower() != "s":
        return

    comando = [
        sys.executable, "-m", "playwright", "codegen",
        "--target", "python", "--output", str(saida), cfg["url_abaco"],
    ]
    print("\n  Gravando... (trabalhe no navegador)\n")
    try:
        subprocess.run(comando, check=True)
    except FileNotFoundError:
        print("  [ERRO] Playwright não instalado. Rode a opção [1] primeiro.")
        pausa()
        return
    except subprocess.CalledProcessError as e:
        print(f"  [ERRO] A gravação terminou com problema: {e}")
        pausa()
        return
    print(f"\n  PRONTO! Fluxo salvo em: {saida}")
    print("  Me mostre o conteúdo desse arquivo para eu montar o robô final.")
    pausa()


# ==========================================================================
# [4] EXTRAIR PARA EXCEL
# ==========================================================================
def extrair(cfg):
    titulo("[4] EXTRAIR PARA EXCEL — lê a tela da Ábaco e grava na planilha")
    try:
        import pandas as pd
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  [ERRO] Bibliotecas não instaladas. Rode a opção [1] primeiro.")
        pausa()
        return

    destino = PASTA / cfg["planilha"]
    print(f"  Os dados lidos serão salvos em: {destino}\n")
    print("  1) faça login na Ábaco;")
    print("  2) navegue até a tela com os dados (aplique filtros se precisar);")
    print("  3) volte aqui e tecle ENTER para ler.\n")
    if input("  Começar? (s/n) ").strip().lower() != "s":
        return

    todas = []
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not cfg["mostrar_navegador"])
        page = navegador.new_page()
        print(f"\n  Abrindo {cfg['url_abaco']}")
        page.goto(cfg["url_abaco"])
        print("\n  >> Faça login e navegue até a tela com os dados.")
        print("     (a senha é digitada no site; não passa por este programa)")

        pagina = 1
        while True:
            input(f"\n  >> Tela pronta? ENTER para LER a página {pagina}...")
            try:
                linhas = ler_tabela_da_tela(page)
            except Exception as e:
                print(f"  [ERRO ao ler a tela] {e}")
                break
            if not linhas:
                print("  Nenhuma tabela encontrada. Se os dados não são uma")
                print("  tabela comum, me mande um print que eu ajusto o leitor.")
            else:
                print(f"  Li {len(linhas)} linhas nesta página.")
                todas.extend(linhas)
            if input("  Ler mais uma página? (avance no site e tecle s / n) "
                     ).strip().lower() != "s":
                break
            pagina += 1
        navegador.close()

    if not todas:
        print("\n  Nada foi extraído.")
        pausa()
        return

    df = pd.DataFrame(todas)
    df.to_excel(destino, index=False)
    print(f"\n  PRONTO! {len(df)} linhas salvas em: {destino}")
    print(f"  Colunas: {', '.join(map(str, df.columns))}")
    pausa()


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
# [5] CONFIGURAÇÕES
# ==========================================================================
def configurar(cfg):
    while True:
        titulo("[5] CONFIGURAÇÕES")
        print(f"  1) URL da Ábaco ....... {cfg['url_abaco']}")
        print(f"  2) Arquivo do fluxo ... {cfg['arquivo_fluxo']}")
        print(f"  3) Planilha ........... {cfg['planilha']}")
        print(f"  4) Nome do relatório .. {cfg['relatorio']}")
        print(f"  5) Mostrar navegador .. {'sim' if cfg['mostrar_navegador'] else 'não'}")
        print("  0) Voltar")
        op = input("\n  O que mudar? ").strip()
        if op == "1":
            cfg["url_abaco"] = input("  Nova URL: ").strip() or cfg["url_abaco"]
        elif op == "2":
            cfg["arquivo_fluxo"] = input("  Arquivo de fluxo: ").strip() or cfg["arquivo_fluxo"]
        elif op == "3":
            cfg["planilha"] = input("  Planilha: ").strip() or cfg["planilha"]
        elif op == "4":
            cfg["relatorio"] = input("  Nome do relatório: ").strip() or cfg["relatorio"]
        elif op == "5":
            cfg["mostrar_navegador"] = not cfg["mostrar_navegador"]
        elif op == "0":
            salvar_config(cfg)
            return
        salvar_config(cfg)


# ==========================================================================
# MENU PRINCIPAL
# ==========================================================================
def menu():
    cfg = carregar_config()
    while True:
        titulo("AUTOMAÇÃO ÁBACO + EXCEL")
        print("  [1] Instalar         (prepara o computador — uma vez)")
        print("  [2] OBSERVAR         (1º passo: documenta como se trabalha)")
        print("  [3] Gravar fluxo     (depois de decidir o que automatizar)")
        print("  [4] Extrair p/ Excel (lê os dados da tela -> planilha)")
        print("  [5] Configurações    (URL, planilha, etc.)")
        print("  [0] Sair")
        op = input("\n  Escolha: ").strip()
        if op == "1":
            instalar()
        elif op == "2":
            observar(cfg)
        elif op == "3":
            gravar(cfg)
        elif op == "4":
            extrair(cfg)
        elif op == "5":
            configurar(cfg)
        elif op == "0":
            print("\n  Até logo!\n")
            return
        else:
            print("  Opção inválida.")


if __name__ == "__main__":
    try:
        menu()
    except KeyboardInterrupt:
        print("\n\n  Encerrado pelo usuário.\n")
