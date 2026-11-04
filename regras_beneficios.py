# -*- coding: utf-8 -*-
"""
==========================================================================
  REGRAS DE NEGÓCIO — Vale-Refeição (VR) e Vale-Alimentação (VA)
==========================================================================

Segue o jeito que a Divisão de Benefícios trabalha hoje (reunião de
03/11/2026, POP de VA/VR revisado em 07/2024 e a planilha "11- NOVEMBRO -
2026.xlsx"). Cada regra está explicada em REGRAS_DE_NEGOCIO.md (os códigos
R1, E1, N2... citados aqui são de lá).

Como a planilha da Divisão funciona, e como este módulo reproduz:

  * A base é a aba CARGA do MÊS ANTERIOR: a lista de servidores com o TIPO
    de cada um (NORMAL, PLANTÃO, DIFERENCIADO, DIFERENCIADO 2). Todo mês a
    lista é atualizada, não refeita.
  * O relatório do SIGA ("situação de afastamento") diz quem continua
    ativo, quem está afastado, quem foi admitido e o salário.
  * QTDE = dias úteis da competência; TOTAL = (QTDE − férias − faltas) ×
    valor do tipo − desconto + reembolso. Plantonista: QTDE = plantões
    informados pela secretaria, sem dedução.
  * VA: crédito de valor fixo e desconto de R$ 1,00; quem tem falta perde
    o mês inteiro (e não tem o desconto).

Este módulo NÃO acessa a Ábaco nem a internet: só lê as planilhas da
pasta de entrada e calcula. Por isso dá para testar sem o sistema
(veja test_regras_beneficios.py).
==========================================================================
"""

import calendar
import io
import re
import unicodedata
from collections import Counter
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

VERSAO_REGRAS = 3
MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
         "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# --------------------------------------------------------------------------
# REGRAS PADRÃO  (copiadas para config.json > "beneficios", onde podem ser
# ajustadas sem mexer no código)
# --------------------------------------------------------------------------
_EXCLUSOES_GERAIS = [
    {"campo": "vinculo", "contem": "estagi",
     "motivo": "Estagiário não recebe (E1)"},
    {"campo": "vinculo", "contem": "aposentad",
     "motivo": "Aposentado não recebe (E1)"},
    {"campo": "vinculo", "contem": "pensionista",
     "motivo": "Pensionista não recebe (E1)"},
    {"campo": "vinculo", "contem": "sentenca",
     "motivo": "Sentença judicial/indenização: não é servidor em exercício (E1)"},
    {"campo": "secretaria", "contem": "inativos",
     "motivo": "Servidor inativo não recebe (E1)"},
]

REGRAS_PADRAO = {
    "versao": VERSAO_REGRAS,
    # {"AAAA-MM-DD": "descrição"} — feriados nacionais já são automáticos
    "feriados_municipais": {},
    "pontos_facultativos": {},
    # D1: servidores da Educação: o VR é feito pela SEMED
    "secretaria_educacao_contem": "educa",
    # E4: tipos de afastamento (SIGA) que tiram o direito no período
    "afastamentos_sem_direito": ["previdencia", "inss", "sem vencimento",
                                 "vacancia", "incapacidade", "invalidez",
                                 "susp"],
    # E4: situações (SIGA) que tiram o servidor da planilha
    "situacoes_excluir": ["exonera", "demitid", "falecid", "desligad",
                          "rescis"],
    # E5: situações que travam a linha até alguém verificar
    "situacoes_verificar": ["aposentacao"],
    "vr": {
        # R7: valor por DIA de cada tipo (aba CARGA, células K2:L5)
        "valor_dia_por_tipo": {"NORMAL": 56.0, "PLANTÃO": 84.0,
                               "DIFERENCIADO": 42.0, "DIFERENCIADO 2": 28.0},
        # R8: quem recebe acima do limite não faz jus ao valor diferenciado
        "limite_salarial": 5662.23,
        "tipos_com_limite_salarial": ["DIFERENCIADO", "DIFERENCIADO 2"],
        # N6: sugestão de tipo para novo admitido pela jornada do cargo
        "horas_para_tipo": {"40": "NORMAL", "30": "DIFERENCIADO",
                            "25": "DIFERENCIADO", "20": "DIFERENCIADO 2",
                            "12": "PLANTÃO"},
        # R6: plantonista novo sem planilha de plantões
        "plantao_padrao": 15,
        "exclusoes": deepcopy(_EXCLUSOES_GERAIS),
        # R9: colunas do arquivo importado na Ábaco: [título, campo]
        "layout_importacao": [["Matrícula", "matricula"],
                              ["Valor", "valor_total"]],
    },
    "va": {
        "ativo": True,
        "valor_credito": 400.0,         # A1: crédito fixo (POP)
        "valor_desconto": 1.0,          # A3: desconto de R$ 1,00 (POP)
        # A2: com falta injustificada: "perde", "revisar" ou "manter"
        "com_falta": "perde",
        # A5: o VA vale para todos os servidores, inclusive a Educação
        # (Lei 3.035/2023, art. 79-A do Estatuto: "independentemente da jornada")
        "inclui_educacao": True,
        # A6: afastamentos que tiram o VA. A Lei 3.035/2023 (art. 79-A, §§ 2º e
        # 3º) mantém o VA no afastamento pelo INSS e só o tira na ausência
        # injustificada e nas licenças sem remuneração.
        "afastamentos_sem_direito": ["sem vencimento", "sem remuneracao",
                                     "vacancia", "incapacidade", "invalidez",
                                     "susp"],
        "exclusoes": deepcopy(_EXCLUSOES_GERAIS) + [
            {"campo": "vinculo", "contem": "politic",
             "motivo": "Agente político não recebe VA (E2)"},
            {"campo": "vinculo", "contem": "proc sel", "ano_admissao": 2022,
             "motivo": "PSS de 2022 não faz jus ao VA (E3)"},
            {"campo": "vinculo", "contem": "pss", "ano_admissao": 2022,
             "motivo": "PSS de 2022 não faz jus ao VA (E3)"},
        ],
        "layout_credito": [["Matrícula", "matricula"], ["Valor", "valor"]],
        "layout_desconto": [["Matrícula", "matricula"], ["Valor", "desconto"]],
    },
    # C1/C2: só para lembrete na tela
    "prazos": {"documentacao_ate_dia": 5, "folha_fecha_dia": 20},
    # nomes de coluna extras, ex.: {"data_admissao": ["dt adm"]}
    "colunas_extras": {},
}

# --------------------------------------------------------------------------
# Nomes de coluna aceitos em cada relatório (comparados sem acento/caixa)
# --------------------------------------------------------------------------
COLUNAS = {
    "matricula": ["matricula", "mat", "matr", "n matricula", "no matricula",
                  "numero matricula", "numero da matricula", "cod servidor",
                  "codigo servidor", "codigo do servidor"],
    "nome": ["nome", "nome servidor", "nome do servidor", "servidor"],
    "cpf": ["cpf", "n cpf", "numero cpf"],
    "cargo": ["cargo", "cargo efetivo", "funcao", "cargo funcao",
              "descricao cargo", "descricao do cargo"],
    "cargo_comissionado": ["cargo comissionado", "cargo em comissao"],
    "jornada": ["jornada", "carga horaria", "ch", "jornada de trabalho",
                "carga horaria semanal"],
    "secretaria": ["secretaria", "nome secretaria", "lotacao", "orgao",
                   "unidade", "local de trabalho", "secretaria lotacao"],
    "vinculo": ["vinculo", "regime", "tipo vinculo", "tipo de vinculo",
                "regime juridico", "natureza do vinculo"],
    "data_admissao": ["data admissao", "data de admissao", "admissao",
                      "nomeacao", "data nomeacao", "data de nomeacao",
                      "dt admissao", "data exercicio", "data de exercicio",
                      "inicio exercicio"],
    "data_exoneracao": ["exoneracao", "data exoneracao", "data de exoneracao",
                        "fim contrato", "termino contrato", "data rescisao"],
    "situacao": ["situacao", "situacao funcional", "status",
                 "situacao de afastamento"],
    "tipo_afastamento": ["tipo afastamento", "tipo de afastamento",
                         "motivo afastamento"],
    "inicio_afastamento": ["inicio afastamento", "inicio do afastamento",
                           "data inicio afastamento"],
    "termino_afastamento": ["termino afastamento", "fim afastamento",
                            "termino do afastamento", "data fim afastamento"],
    "salario": ["salario", "salario base", "vencimento base", "vencimento"],
    "tipo": ["tipo", "tipo vr", "tipo de vale", "tipo beneficio"],
    "periodo": ["periodo de gozo", "periodo gozo", "periodo",
                "periodo de ferias"],
    "competencia": ["competencia"],
    "data": ["data", "data falta", "data da falta", "dia"],
    "quantidade": ["quantidade", "qtd", "qtde", "dias", "plantoes",
                   "qtd plantoes", "quantidade de plantoes",
                   "quantidade plantoes", "n de plantoes", "no de plantoes",
                   "n plantoes", "no plantoes", "numero de plantoes",
                   "faltas", "qtd faltas"],
    "inicio": ["inicio", "data inicio", "data de inicio", "inicio gozo",
               "dt inicio", "inicio ferias", "inicio do gozo"],
    "fim": ["fim", "data fim", "data de fim", "fim gozo", "dt fim",
            "termino", "fim ferias", "fim do gozo", "data termino"],
    "desconto": ["desconto", "desconto r"],
    "reembolso": ["reembolso", "ressarcimento", "reembolso r"],
    "observacao": ["observacao", "observacoes", "obs"],
}

# Como os campos aparecem nas planilhas geradas
ROTULOS = {"matricula": "Matrícula", "nome": "Nome", "cargo": "Cargo",
           "jornada": "Jornada", "secretaria": "Secretaria",
           "vinculo": "Vínculo", "data_admissao": "Data de admissão",
           "situacao": "Situação", "data": "Data", "inicio": "Início",
           "fim": "Fim", "tipo": "Tipo", "quantidade": "Quantidade",
           "valor_dia": "Valor diário", "valor_total": "Valor total",
           "valor": "Crédito", "desconto": "Desconto"}

# Modelos de planilha de entrada (botão "Gerar modelos vazios")
MODELOS = {
    "servidores": ["MATRICULA", "NOME", "CPF", "SITUACAO", "REGIME",
                   "NOME_SECRETARIA", "NOMEACAO", "EXONERACAO",
                   "CARGO_EFETIVO", "CARGO_COMISSIONADO", "SALARIO",
                   "INICIO_AFASTAMENTO", "TERMINO_AFASTAMENTO",
                   "TIPO_AFASTAMENTO"],
    "ferias": ["Matricula", "Nome do Servidor", "Tipo de Folha",
               "Período de Gozo"],
    "faltas": ["Nome", "Matricula", "Cargo", "Admissão", "Competência",
               "Qtde"],
    "plantoes": ["Matrícula", "Nome", "Quantidade de plantões"],
    "ajustes": ["Matrícula", "Nome", "Tipo", "Quantidade", "Desconto",
                "Reembolso", "Observação"],
}


class ErroEntrada(Exception):
    """Problema nas planilhas de entrada, com mensagem para quem opera."""


def regras_completas(extra):
    """REGRAS_PADRAO com o que veio do config.json por cima.

    Um config.json da versão anterior (sem "versao") tinha outra estrutura
    (valor por jornada, exclusões únicas): dele só se aproveitam calendário,
    Educação, prazos e colunas."""
    def mesclar(base, novo):
        for k, v in (novo or {}).items():
            if isinstance(v, dict) and isinstance(base.get(k), dict):
                mesclar(base[k], v)
            else:
                base[k] = v
        return base
    extra = deepcopy(extra or {})
    versao = int(extra.get("versao") or 1) if extra else VERSAO_REGRAS
    if versao < 2:
        extra = {k: extra[k] for k in ("feriados_municipais", "pontos_facultativos",
                                       "secretaria_educacao_contem", "prazos",
                                       "colunas_extras") if k in extra}
    elif versao < 3 and isinstance(extra.get("va"), dict):
        # versão 3: cedido com ônus do Município recebe o VA (Lei 3.035/2023)
        va = extra["va"]
        if isinstance(va.get("exclusoes"), list):
            va["exclusoes"] = [e for e in va["exclusoes"]
                               if norm(e.get("contem")) != "cedid"]
    regras = mesclar(deepcopy(REGRAS_PADRAO), extra)
    regras["versao"] = VERSAO_REGRAS
    return regras


# ==========================================================================
# Calendário (R1)
# ==========================================================================
def pascoa(ano):
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return date(ano, mes, dia)


def feriados_nacionais(ano):
    p = pascoa(ano)
    f = {
        date(ano, 1, 1): "Confraternização Universal",
        p - timedelta(days=2): "Sexta-feira Santa",
        date(ano, 4, 21): "Tiradentes",
        date(ano, 5, 1): "Dia do Trabalho",
        date(ano, 9, 7): "Independência",
        date(ano, 10, 12): "Nossa Senhora Aparecida",
        date(ano, 11, 2): "Finados",
        date(ano, 11, 15): "Proclamação da República",
        date(ano, 12, 25): "Natal",
    }
    if ano >= 2024:
        f[date(ano, 11, 20)] = "Consciência Negra"
    return f


def pontos_facultativos_comuns(ano):
    """Não são descontados sozinhos: dependem do decreto do município."""
    p = pascoa(ano)
    return {
        p - timedelta(days=48): "Carnaval (segunda)",
        p - timedelta(days=47): "Carnaval (terça)",
        p - timedelta(days=46): "Quarta-feira de Cinzas",
        p + timedelta(days=60): "Corpus Christi",
    }


def ultimo_dia(ano, mes):
    return date(ano, mes, calendar.monthrange(ano, mes)[1])


def mes_antes(ano, mes, n=1):
    total = ano * 12 + (mes - 1) - n
    return total // 12, total % 12 + 1


@dataclass
class Calendario:
    folgas: dict  # date -> descrição

    @classmethod
    def do_config(cls, ano_ini, ano_fim, regras):
        folgas = {}
        for ano in range(ano_ini, ano_fim + 1):
            folgas.update(feriados_nacionais(ano))
        for chave in ("feriados_municipais", "pontos_facultativos"):
            for txt, desc in (regras.get(chave) or {}).items():
                d = parse_data(txt)
                if d is None:
                    raise ErroEntrada(f"Data inválida em config.json > "
                                      f"{chave}: '{txt}' (use AAAA-MM-DD)")
                folgas[d] = desc
        return cls(folgas)

    def dias_uteis(self, ini, fim):
        """Dias úteis entre ini e fim, inclusive."""
        n, d = 0, ini
        while d <= fim:
            if d.weekday() < 5 and d not in self.folgas:
                n += 1
            d += timedelta(days=1)
        return n

    def dias_uteis_mes(self, ano, mes):
        return self.dias_uteis(date(ano, mes, 1), ultimo_dia(ano, mes))

    def folgas_no_mes(self, ano, mes):
        return sorted((d, desc) for d, desc in self.folgas.items()
                      if d.year == ano and d.month == mes and d.weekday() < 5)


# ==========================================================================
# Leitura das planilhas de entrada
# ==========================================================================
def norm(txt):
    """Texto sem acento, minúsculo, só letras/números separados por espaço."""
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return ""
    s = unicodedata.normalize("NFKD", str(txt))
    s = s.encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return "" if s in ("nan", "none", "nat") else s


def _txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat") or re.fullmatch(r"-+", s):
        return ""
    return s


def matricula_texto(v):
    s = _txt(v)
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def chave_matricula(v):
    """Chave para cruzar planilhas: só dígitos, sem zeros à esquerda."""
    s = matricula_texto(v)
    dig = re.sub(r"\D", "", s)
    return (dig.lstrip("0") or "0") if dig else s.upper()


def parse_data(v):
    if isinstance(v, (pd.Timestamp, date)):
        return None if pd.isna(v) else pd.Timestamp(v).date()
    s = _txt(v)
    if not s:
        return None
    if re.match(r"\d{4}-\d{2}-\d{2}", s):
        d = pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce")
    elif re.fullmatch(r"\d{4,5}(\.0+)?", s):          # número de série do Excel
        d = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(float(s)))
    else:
        d = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.date()


def datas_no_texto(v):
    """'23/11/2026 à 02/12/2026' -> [date(2026, 11, 23), date(2026, 12, 2)]"""
    achadas = re.findall(r"\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2}", _txt(v))
    return [d for d in (parse_data(a) for a in achadas) if d]


def numero(v):
    """'2.681,37', '2681,37', '2681.37', 3.6 -> float (ou None)."""
    if isinstance(v, (int, float)) and not (isinstance(v, float) and pd.isna(v)):
        return float(v)
    s = _txt(v).replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def chave_jornada(v):
    """'4:30', '4h30', '4,5 horas' -> '4h30'; '6 horas' -> '6h'; '40' -> '40h'."""
    s = norm(v)
    if not s:
        return ""
    m = re.match(r"(\d+)\s*(?:h|horas?|hs)?\s*(\d{1,2})?", s)
    dec = re.match(r"(\d+)[.,](\d+)", _txt(v))
    if dec:
        horas = int(dec.group(1))
        minutos = round(float("0." + dec.group(2)) * 60)
    elif m:
        horas, minutos = int(m.group(1)), int(m.group(2) or 0)
    else:
        return s
    return f"{horas}h{minutos:02d}" if minutos else f"{horas}h"


def horas_do_texto(v):
    """'2631 - VIGIA MUNICIPAL - 40 HORAS' -> '40'; '... -30HS' -> '30'."""
    m = re.search(r"(\d{1,2})\s*(?:H\b|HS\b|HORA)", _txt(v).upper())
    return m.group(1) if m else ""


def _aliases(regras):
    aliases = {campo: [norm(a) for a in nomes] for campo, nomes in COLUNAS.items()}
    for campo, nomes in (regras.get("colunas_extras") or {}).items():
        aliases.setdefault(campo, []).extend(norm(a) for a in nomes)
    return aliases


def _mapear(linha, aliases):
    """{posição da coluna: campo} para os cabeçalhos reconhecidos."""
    mapa, usados = {}, set()
    for pos, cel in enumerate(linha):
        n = norm(cel)
        for campo, nomes in aliases.items():
            if campo not in usados and n in nomes:
                mapa[pos] = campo
                usados.add(campo)
                break
    return mapa


def encontrar_arquivo(pasta, nome):
    """O arquivo mais recente entre nome.xlsx / .xls / .csv / .xlsm."""
    achados = [Path(pasta) / f"{nome}{ext}" for ext in (".xlsx", ".xls", ".csv", ".xlsm")]
    achados = [p for p in achados if p.exists()]
    return max(achados, key=lambda p: p.stat().st_mtime) if achados else None


def ler_planilha(caminho):
    """Lê sem cabeçalho (o cabeçalho é localizado depois)."""
    ext = caminho.suffix.lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(caminho, header=None, dtype=str, sep=None,
                                   engine="python", encoding=enc)
            except UnicodeDecodeError:
                continue
    try:
        return pd.read_excel(caminho, header=None, dtype=str)
    except Exception as erro_excel:
        # alguns sistemas exportam uma página HTML com extensão .xls
        bruto = caminho.read_bytes()
        texto = bruto.decode("cp1252", errors="replace")
        for enc in ("utf-8", "cp1252"):
            try:
                texto = bruto.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        try:
            t = max(pd.read_html(io.StringIO(texto)), key=len)
        except Exception:
            raise ErroEntrada(f"Não consegui ler {caminho.name}: {erro_excel}")
        cab = pd.DataFrame([[str(c) for c in t.columns]])
        return pd.concat([cab, pd.DataFrame(t.astype(str).values)],
                         ignore_index=True)


def _tabela(bruto, obrigatorias, aliases, nome_arq):
    """Acha a linha de cabeçalho (relatórios têm título antes) e devolve o
    DataFrame só com as colunas reconhecidas."""
    melhor_i, melhor = 0, {}
    for i in range(min(30, len(bruto))):
        mapa = _mapear(bruto.iloc[i].tolist(), aliases)
        if len(mapa) > len(melhor):
            melhor_i, melhor = i, mapa
    faltando = [c for c in obrigatorias if c not in melhor.values()]
    if faltando:
        lidas = [_txt(c) for c in bruto.iloc[melhor_i].tolist() if _txt(c)] \
            if len(bruto) else []
        raise ErroEntrada(
            f"{nome_arq}: não encontrei a(s) coluna(s) "
            f"{', '.join(ROTULOS.get(c, c) for c in faltando)}.\n"
            f"     Colunas encontradas: {', '.join(lidas) or '(nenhuma)'}\n"
            f"     Se o nome for outro, cadastre em Configurações > Colunas.")
    df = bruto.iloc[melhor_i + 1:, list(melhor)].copy()
    df.columns = list(melhor.values())
    df = df.dropna(how="all")
    df["matricula"] = df["matricula"].map(matricula_texto)
    df = df[df["matricula"] != ""].copy()
    df = df[df["matricula"].map(lambda m: bool(re.search(r"\d", m)))].copy()
    df["chave"] = df["matricula"].map(chave_matricula)
    return df.reset_index(drop=True)


def carregar(pasta, nome, obrigatorias, aliases, opcional=False):
    arq = encontrar_arquivo(Path(pasta), nome)
    if arq is None:
        if opcional:
            return None
        raise ErroEntrada(f"Falta o arquivo '{nome}.xlsx' na pasta {pasta}. "
                          f"Use Benefícios VR/VA > Coletar relatórios ou "
                          f"adicione o arquivo manualmente.")
    return _tabela(ler_planilha(arq), obrigatorias, aliases, arq.name)


# ---- planilha de carga do mês anterior (a base do mês)
@dataclass
class LinhaBase:
    matricula: str
    chave: str
    nome: str
    tipo: str
    qtde_fixa: float     # número digitado em QTDE; None = fórmula/padrão
    observacao: str
    ordem: int
    qtde_valor: float = None   # valor que a QTDE mostrava (fórmula calculada)
    total_fixo: float = None   # TOTAL digitado à mão (sem fórmula)


def _cabecalho_carga(valores):
    n = [norm(v) for v in valores]
    return ("mat" in n or "matricula" in n) and "tipo" in n


def ler_carga_anterior(caminho, aliases=None):
    """Lê a aba CARGA da planilha do mês anterior. Para .xlsx/.xlsm lê as
    fórmulas: QTDE com fórmula (=$I$3) segue o padrão do mês; QTDE com
    número foi digitada à mão (plantões, zerados...)."""
    caminho = Path(caminho)
    aliases = aliases or _aliases({})
    linhas_brutas, calculadas = [], []
    if caminho.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(caminho, read_only=True, data_only=False)
        except Exception as e:
            raise ErroEntrada(f"Não consegui abrir {caminho.name}: {e}")
        nomes = sorted(wb.sheetnames, key=lambda s: norm(s) != "carga")
        for nome in nomes:
            linhas = list(wb[nome].iter_rows(max_row=40, values_only=True))
            if any(_cabecalho_carga(l) for l in linhas):
                linhas_brutas = [list(l) for l in wb[nome].iter_rows(values_only=True)]
                # o valor que cada fórmula mostrava na última vez que o Excel salvou
                wv = load_workbook(caminho, read_only=True, data_only=True)
                calculadas = [list(l) for l in wv[nome].iter_rows(values_only=True)]
                wv.close()
                break
        wb.close()
    else:
        bruto = ler_planilha(caminho)
        linhas_brutas = bruto.where(pd.notna(bruto), None).values.tolist()
        calculadas = linhas_brutas
    ini = next((i for i, l in enumerate(linhas_brutas[:40]) if _cabecalho_carga(l)), None)
    if ini is None:
        raise ErroEntrada(f"{caminho.name}: não encontrei a aba CARGA (cabeçalho "
                          f"com MAT e TIPO). Use a planilha de carga do mês "
                          f"anterior.")
    mapa = _mapear(linhas_brutas[ini], aliases)
    pos = {campo: p for p, campo in mapa.items()}
    cab = [norm(v) for v in linhas_brutas[ini]]
    if "total" in cab:
        pos["_total"] = cab.index("total")
    base, vistos, repetidas = [], set(), []
    for ordem, l in enumerate(linhas_brutas[ini + 1:]):
        pega = lambda campo, l=l: l[pos[campo]] \
            if campo in pos and pos[campo] < len(l) else None
        lc = calculadas[ini + 1 + ordem] if ini + 1 + ordem < len(calculadas) else []
        mat = matricula_texto(pega("matricula"))
        if not mat or not re.search(r"\d", mat):
            continue
        chave = chave_matricula(mat)
        if chave in vistos:
            repetidas.append(mat)
            continue
        vistos.add(chave)
        q = pega("quantidade")
        fixa = None
        if isinstance(q, (int, float)) and not isinstance(q, bool):
            fixa = float(q)
        elif isinstance(q, str) and not q.strip().startswith("="):
            fixa = numero(q)
        valor = numero(pega("quantidade", lc)) if fixa is None else fixa
        t = pega("_total")
        total_fixo = float(t) if isinstance(t, (int, float)) and \
            not isinstance(t, bool) else None
        base.append(LinhaBase(mat, chave, _txt(pega("nome")), _txt(pega("tipo")),
                              fixa, _txt(pega("observacao")), ordem, valor,
                              total_fixo))
    if not base:
        raise ErroEntrada(f"{caminho.name}: a aba CARGA está vazia.")
    return base, repetidas


# ==========================================================================
# Cálculo
# ==========================================================================
COLS_INFO = ["Matrícula", "Nome", "Cargo", "Secretaria", "Vínculo",
             "Admissão", "Situação"]
COLS_VR = COLS_INFO + ["Origem", "Tipo", "Dias úteis", "Afastamento",
                       "Faltas", "Férias (dias úteis)", "Total deduções",
                       "Quantidade", "Valor diário", "Desconto", "Reembolso",
                       "Valor total", "Observação", "Obs. mês anterior",
                       "Status"]
COLS_VA = COLS_INFO + ["Faltas", "Crédito", "Desconto", "Observação", "Status"]
COLS_COMP = COLS_INFO + ["Tipo", "Dias retroativos (mês da admissão)",
                         "Dias do mês seguinte", "Total de dias",
                         "Valor diário", "Valor total", "Status"]
COLS_PEND = ["Matrícula", "Nome", "Benefício", "Motivo", "O que fazer",
             "Trava a importação"]


@dataclass
class Resultado:
    ano: int
    mes: int
    dias_uteis: int
    folgas: list
    vr: pd.DataFrame
    va: pd.DataFrame
    complementar: pd.DataFrame
    excluidos: pd.DataFrame
    educacao: pd.DataFrame
    aguardando: pd.DataFrame
    pendencias: pd.DataFrame
    fora: pd.DataFrame = None
    avisos: list = field(default_factory=list)
    valores_tipo: dict = field(default_factory=dict)
    ferias_usadas: pd.DataFrame = None
    faltas_usadas: pd.DataFrame = None
    carga_anterior: str = ""

    @property
    def rotulo(self):
        return f"{self.ano}-{self.mes:02d}"

    @property
    def folha(self):
        """(ano, mês) da folha em que o crédito é lançado: o mês anterior."""
        return mes_antes(self.ano, self.mes, 1)


def _contem(valor, trechos):
    v = norm(valor)
    return bool(v) and any(norm(t) and norm(t) in v for t in trechos)


def _motivo_exclusao(r, adm, regras_lista):
    for regra in regras_lista:
        if not _contem(r.get(regra["campo"]), [regra["contem"]]):
            continue
        ano = regra.get("ano_admissao")
        if ano and (adm is None or adm.year != int(ano)):
            continue
        return regra["motivo"]
    return None


def _eh_educacao(r, regras):
    trecho = regras.get("secretaria_educacao_contem")
    return bool(trecho) and _contem(r.get("secretaria"), [trecho])


def _iguais(campo, a, b):
    if campo == "data_admissao":
        return parse_data(a) == parse_data(b)
    if campo == "jornada":
        return chave_jornada(a) == chave_jornada(b)
    return norm(a) == norm(b)


def _cargo(r):
    ef, cc = _txt(r.get("cargo")), _txt(r.get("cargo_comissionado"))
    return " / ".join(c for c in (ef, cc) if c)


def _info(r):
    adm = parse_data(r.get("data_admissao"))
    return {"Matrícula": r["matricula"], "Nome": _txt(r.get("nome")),
            "Cargo": _cargo(r), "Secretaria": _txt(r.get("secretaria")),
            "Vínculo": _txt(r.get("vinculo")),
            "Admissão": adm.strftime("%d/%m/%Y") if adm else "",
            "Situação": _txt(r.get("situacao"))}


def _tipo_valido(txt, valores):
    """Normaliza 'plantao', 'Diferenciado 2'... para a chave do config."""
    n = norm(txt)
    for t in valores:
        if norm(t) == n:
            return t
    return ""


def _br(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _servidores(pasta, aliases):
    return carregar(pasta, "servidores", ["matricula", "nome", "data_admissao"],
                    aliases)


def _arquivo_carga(pasta, carga_anterior):
    arq = Path(carga_anterior) if carga_anterior else \
        encontrar_arquivo(Path(pasta), "carga_anterior")
    if arq is None or not arq.exists():
        raise ErroEntrada(
            "Falta a planilha de carga do MÊS ANTERIOR (a base do cálculo, com "
            "o TIPO de cada servidor). Adicione a planilha do mês passado "
            "(ex.: '11- NOVEMBRO - 2026.xlsx') como 'Planilha do mês anterior'. "
            "Nos meses seguintes o programa usa a que ele mesmo gerou.")
    return arq


def novos_admitidos(pasta, ano, mes, regras, carga_anterior=None):
    """Admitidos a partir do mês M-2, fora da Educação e que não estão na
    planilha do mês anterior: são os contratos a conferir na Ábaco (N1)."""
    aliases = _aliases(regras)
    serv = _servidores(pasta, aliases)
    try:
        na_base = {b.chave for b in ler_carga_anterior(
            _arquivo_carga(pasta, carga_anterior), aliases)[0]}
    except ErroEntrada:
        na_base = set()
    a2, m2 = mes_antes(ano, mes, 2)
    ini = date(a2, m2, 1)
    novos = []
    for r in serv.to_dict("records"):
        adm = parse_data(r.get("data_admissao"))
        if adm and adm >= ini and r["chave"] not in na_base \
                and not _eh_educacao(r, regras):
            novos.append((r["matricula"], _txt(r.get("nome"))))
    return novos


def calcular(pasta_entrada, ano, mes, regras, dias_uteis=None,
             carga_anterior=None):
    """Calcula VR, VA e complementar da competência (mês do benefício).
    A folha (crédito) é a do mês anterior: competência 12/2026 = folha de
    11/2026, como na planilha '11- NOVEMBRO - 2026'."""
    pasta = Path(pasta_entrada)
    aliases = _aliases(regras)
    avisos = []
    pendencias = []

    def pend(info, beneficio, motivo, acao, trava=True):
        pendencias.append({"Matrícula": info["Matrícula"], "Nome": info["Nome"],
                           "Benefício": beneficio, "Motivo": motivo,
                           "O que fazer": acao,
                           "Trava a importação": "Sim" if trava else "Não"})

    # ---------------------------------------------------------- entradas
    serv = _servidores(pasta, aliases)
    arq_carga = _arquivo_carga(pasta, carga_anterior)
    base, repetidas = ler_carga_anterior(arq_carga, aliases)
    avisos.append(f"Base: {arq_carga.name} ({len(base)} servidores na aba CARGA).")
    for mat in repetidas:
        pend({"Matrícula": mat, "Nome": ""}, "VR",
             "Matrícula repetida na planilha do mês anterior",
             "Foi usada só a primeira linha.", False)
    faltas = carregar(pasta, "faltas", ["matricula"], aliases, opcional=True)
    ferias = carregar(pasta, "ferias", ["matricula"], aliases, opcional=True)
    plantoes = carregar(pasta, "plantoes", ["matricula", "quantidade"], aliases,
                        opcional=True)
    conf = carregar(pasta, "conferencia_abaco", ["matricula"], aliases,
                    opcional=True)
    ajustes = carregar(pasta, "ajustes", ["matricula"], aliases, opcional=True)

    cal = Calendario.do_config(ano - 1, ano + 1, regras)
    du_calc = cal.dias_uteis_mes(ano, mes)
    du = du_calc if dias_uteis is None else int(dias_uteis)
    if du != du_calc:
        avisos.append(f"Dias úteis informados manualmente: {du} "
                      f"(o calendário calculou {du_calc}).")
    ini_comp, fim_comp = date(ano, mes, 1), ultimo_dia(ano, mes)
    a1, m1 = mes_antes(ano, mes, 1)
    a2, m2 = mes_antes(ano, mes, 2)
    ini_c1, ini_c2, fim_c2 = date(a1, m1, 1), date(a2, m2, 1), ultimo_dia(a2, m2)
    du_c1 = cal.dias_uteis_mes(a1, m1)

    vr_cfg, va_cfg = regras["vr"], regras["va"]
    valores = {t: (None if v in (None, "") else float(v))
               for t, v in vr_cfg["valor_dia_por_tipo"].items()}
    if not any(v is not None for v in valores.values()):
        avisos.append("Valores diários do VR não configurados: as planilhas "
                      "saem só com quantidades.")
    limite = numero(vr_cfg.get("limite_salarial"))
    com_limite = {_tipo_valido(t, valores) for t in
                  vr_cfg.get("tipos_com_limite_salarial") or []}

    # ---------------------------------------------------------- SIGA
    for col in ("nome", "cpf", "cargo", "cargo_comissionado", "jornada",
                "secretaria", "vinculo", "situacao", "data_exoneracao",
                "tipo_afastamento", "inicio_afastamento", "termino_afastamento",
                "salario"):
        if col not in serv:
            serv[col] = None
    for r in serv[serv.duplicated("chave", keep="first")].to_dict("records"):
        pend(_info(r), "VR/VA", "Matrícula repetida no relatório do SIGA",
             "Foi usada só a primeira linha. Confira o relatório.", False)
    serv = serv.drop_duplicates("chave", keep="first")
    siga = {r["chave"]: r for r in serv.to_dict("records")}

    # conferência na Ábaco (passo 2): o valor da Ábaco prevalece (N1)
    jornada_abaco = {}
    if conf is not None:
        n_conf = 0
        for c in conf.to_dict("records"):
            r = siga.get(c["chave"])
            if r is None:
                continue
            n_conf += 1
            if _txt(c.get("jornada")):
                jornada_abaco[c["chave"]] = _txt(c.get("jornada"))
            for campo in ("nome", "data_admissao", "cargo", "secretaria"):
                novo = _txt(c.get(campo))
                if not novo or _iguais(campo, r.get(campo), novo):
                    continue
                antigo = _txt(r.get(campo))
                r[campo] = novo
                pend(_info(r), "VR/VA", f"{ROTULOS[campo]} diferente: SIGA "
                                        f"'{antigo}' x Ábaco '{novo}'",
                     "Usado o valor da Ábaco. Confira no contrato.", False)
        avisos.append(f"Conferência da Ábaco aplicada ({n_conf} matrícula(s)).")

    # ---------------------------------------------------------- faltas (R2/A2)
    faltas_por, faltas_usadas = {}, None
    if faltas is None:
        avisos.append("Relatório de faltas não encontrado: nenhuma falta "
                      "foi descontada.")
    else:
        if "quantidade" in faltas:
            faltas["qtd"] = faltas["quantidade"].map(numero).fillna(1)
        elif "data" in faltas:
            faltas["qtd"] = 1.0
            faltas = faltas.drop_duplicates(["chave", "data"])
        else:
            faltas["qtd"] = 1.0
        faltas_por = faltas.groupby("chave")["qtd"].sum().to_dict()
        if "competencia" in faltas:
            comps = sorted({d.strftime("%m/%Y") for d in
                            faltas["competencia"].map(parse_data) if d})
            if comps:
                avisos.append(f"Faltas: relatório com competência(s) "
                              f"{', '.join(comps)} (todas as linhas foram "
                              f"consideradas).")
        faltas_usadas = faltas.drop(columns=["chave"])

    # ---------------------------------------------------------- férias (R2-R4)
    ferias_por, ferias_linhas = {}, []
    if ferias is None:
        avisos.append("Relatório de férias não encontrado: nenhuma férias "
                      "foi descontada.")
    else:
        if "periodo" not in ferias and not {"inicio", "fim"} <= set(ferias):
            raise ErroEntrada("Relatório de férias: não encontrei a coluna "
                              "'Período de Gozo' nem as colunas Início e Fim.")
        invalidas = 0
        for r in ferias.to_dict("records"):
            if "periodo" in r and _txt(r.get("periodo")):
                datas = datas_no_texto(r.get("periodo"))
                ini, fim = (datas[0], datas[-1]) if len(datas) >= 2 else (None, None)
            else:
                ini, fim = parse_data(r.get("inicio")), parse_data(r.get("fim"))
            if ini is None or fim is None or fim < ini:
                invalidas += 1
                continue
            ferias_por.setdefault(r["chave"], []).append((ini, fim))
            if ini <= fim_comp and fim >= ini_comp:
                i2, f2 = max(ini, ini_comp), min(fim, fim_comp)
                ferias_linhas.append({
                    "Matrícula": r["matricula"], "Nome": _txt(r.get("nome")),
                    "Período de gozo": f"{ini:%d/%m/%Y} à {fim:%d/%m/%Y}",
                    "Início ajustado": i2, "Fim ajustado": f2,
                    "Dias úteis": cal.dias_uteis(i2, f2)})
        if invalidas:
            avisos.append(f"{invalidas} linha(s) do relatório de férias com "
                          f"período inválido foram ignoradas.")

    # ---------------------------------------------------------- plantões (R6)
    plantoes_por = {}
    if plantoes is None:
        avisos.append("Planilha de plantões não encontrada: os plantonistas "
                      "ficam com a quantidade do mês anterior (confira).")
    else:
        q = plantoes["quantidade"].map(numero)
        plantoes_por = {k: int(v) for k, v in
                        q.groupby(plantoes["chave"]).sum().items()}

    # ---------------------------------------------------------- ajustes
    ajustes_por = {}
    if ajustes is not None:
        for r in ajustes.to_dict("records"):
            a = ajustes_por.setdefault(r["chave"], {})
            t = _txt(r.get("tipo"))
            if t:
                tv = _tipo_valido(t, valores)
                if not tv:
                    raise ErroEntrada(f"ajustes: tipo '{t}' da matrícula "
                                      f"{r['matricula']} não existe. Use: "
                                      f"{', '.join(valores)}.")
                a["tipo"] = tv
            for campo in ("quantidade", "desconto", "reembolso"):
                v = numero(r.get(campo))
                if v is not None:
                    a[campo] = v
            if _txt(r.get("observacao")):
                a["observacao"] = _txt(r.get("observacao"))
        avisos.append(f"Ajustes manuais aplicados: {len(ajustes_por)} matrícula(s).")

    # ---------------------------------------------------------- apoio
    def ded_ferias(chave):
        return min(du, sum(cal.dias_uteis(max(i, ini_comp), min(f, fim_comp))
                           for i, f in ferias_por.get(chave, [])
                           if i <= fim_comp and f >= ini_comp))

    def ded_faltas(chave):
        return int(min(faltas_por.get(chave, 0), du))

    def afastamento(r, sem_direito=None):
        """(sem_direito, dias úteis afastados no mês, mês inteiro, texto)"""
        tipo = _txt(r.get("tipo_afastamento"))
        if not tipo and not _contem(r.get("situacao"), ["afastad"]):
            return False, 0, False, ""
        ini = parse_data(r.get("inicio_afastamento"))
        fim = parse_data(r.get("termino_afastamento"))
        desde = f" desde {ini:%d/%m/%Y}" if ini else ""
        ate = f" até {fim:%d/%m/%Y}" if fim else ""
        texto = f"{tipo or 'afastado (sem tipo no SIGA)'}{desde}{ate}"
        if (ini and ini > fim_comp) or (fim and fim < ini_comp):
            return False, 0, False, ""
        sem = _contem(tipo, regras["afastamentos_sem_direito"]
                      if sem_direito is None else sem_direito)
        i2, f2 = max(ini or ini_comp, ini_comp), min(fim or fim_comp, fim_comp)
        inteiro = i2 <= ini_comp and f2 >= fim_comp
        return sem, cal.dias_uteis(i2, f2), inteiro, texto

    def salario(r):
        return numero(r.get("salario"))

    # tipo mais comum de cada cargo na planilha anterior (sugestão para novos)
    por_cargo = {}
    for b in base:
        r = siga.get(b.chave)
        t = _tipo_valido(b.tipo, valores)
        if r is not None and t:
            por_cargo.setdefault(norm(_cargo(r)), Counter())[t] += 1

    def sugerir_tipo(chave, r):
        """(tipo, explicação, trava)"""
        if _contem(r.get("vinculo"), ["comission"]) or \
                (_txt(r.get("cargo_comissionado")) and not _txt(r.get("cargo"))):
            return "NORMAL" if "NORMAL" in valores else "", "cargo comissionado", False
        cont = por_cargo.get(norm(_cargo(r)))
        if cont:
            t, n = cont.most_common(1)[0]
            total = sum(cont.values())
            if total >= 3 and n / total >= 0.9:
                return t, f"{n} de {total} servidores com esse cargo são {t}", False
        m = re.match(r"\d+", chave_jornada(jornada_abaco.get(chave, "")))
        horas = m.group(0) if m else horas_do_texto(_cargo(r))
        t = _tipo_valido((vr_cfg.get("horas_para_tipo") or {}).get(horas, ""), valores)
        if t:
            return t, f"jornada de {horas} horas", True
        return "", "sem jornada no cargo", True

    # QTDE digitada à mão mais comum (ex.: os dias úteis do mês anterior)
    fixas = Counter(b.qtde_fixa for b in base if b.qtde_fixa
                    and _tipo_valido(b.tipo, valores) != "PLANTÃO")
    qtde_comum = fixas.most_common(1)[0][0] if fixas else None

    # ---------------------------------------------------------- VR
    vr, comp, excl, edu, aguard, fora = [], [], [], [], [], []
    cpf_vr = {}

    def linha_vr(chave, r, b, origem):
        info = _info(r)
        aj = ajustes_por.get(chave, {})
        obs, travas = [], []
        obs_ant = b.observacao if b else ""

        # tipo
        if aj.get("tipo"):
            tipo = aj["tipo"]
        elif b is not None:
            tipo = _tipo_valido(b.tipo, valores)
            if not tipo:
                pend(info, "VR", f"Tipo '{b.tipo}' da planilha anterior não "
                                 f"existe na configuração",
                     f"Informe o tipo em ajustes ({', '.join(valores)}).")
                travas.append("tipo")
        else:
            tipo, porque, trava = sugerir_tipo(chave, r)
            if trava or not tipo:
                pend(info, "VR", f"Novo admitido: confirme o tipo no contrato "
                                 f"(sugerido: {tipo or '—'}, {porque})",
                     "Informe o Tipo da matrícula no arquivo de ajustes.")
                travas.append("tipo")
            else:
                obs.append(f"Tipo {tipo}: {porque}")
                if porque != "cargo comissionado":
                    pend(info, "VR", f"Novo admitido: tipo {tipo} sugerido "
                                     f"({porque})",
                         "Confira no contrato; se for outro, informe em ajustes.",
                         False)
        valor_dia = valores.get(tipo)
        if tipo and valor_dia is None:
            pend(info, "VR", f"Tipo {tipo} sem valor diário configurado",
                 "Configurações > Valores.")
            travas.append("valor")

        # base do mês: dias úteis, ou até o fim do contrato
        exon = parse_data(r.get("data_exoneracao"))
        base_dias = du
        if exon and ini_comp <= exon < fim_comp:
            base_dias = cal.dias_uteis(ini_comp, exon)
            obs.append(f"Contrato/exoneração em {exon:%d/%m/%Y}: proporcional")
            pend(info, "VR", f"Termina em {exon:%d/%m/%Y}, dentro da competência",
                 "Quantidade proporcional até essa data. Confira.", False)

        sem_dir, dias_af, inteiro, txt_af = afastamento(r)
        sal = salario(r)
        zero_regra = ""
        if tipo in com_limite and limite and sal and sal > limite:
            zero_regra = f"NÃO FAZ JUS: recebe {_br(sal)}, acima de {_br(limite)}"
        elif sem_dir and inteiro:
            zero_regra = f"AFASTADO: {txt_af}"

        fer = fal = af = 0
        if tipo == "PLANTÃO":
            if "quantidade" in aj:
                qtde = aj["quantidade"]
            elif chave in plantoes_por:
                qtde = plantoes_por[chave]
            elif b is not None and b.qtde_valor is not None:
                qtde = b.qtde_valor
                pend(info, "VR", f"Plantão não informado pela secretaria: "
                                 f"mantida a quantidade do mês anterior ({qtde:g})",
                     "Confira; se mudou, informe em ajustes.", False)
            else:
                qtde = int(vr_cfg.get("plantao_padrao") or 0)
                pend(info, "VR", "Plantonista sem quantidade de plantões",
                     "Confira a planilha da secretaria ou informe em ajustes.")
                travas.append("plantao")
            if zero_regra and "quantidade" not in aj:
                qtde = 0
            elif sem_dir and dias_af:
                pend(info, "VR", f"Plantonista afastado parte do mês ({txt_af})",
                     "Confira os plantões.", False)
            deducoes = 0
        else:
            zerado_antes = b is not None and b.qtde_fixa == 0
            pela_regra = zerado_antes and b.observacao.startswith(
                ("AFASTADO:", "NÃO FAZ JUS:"))       # zerado pelo programa
            carrega_zero = (zerado_antes and not pela_regra
                            and "ferias" not in norm(b.observacao))
            if pela_regra and not zero_regra and "quantidade" not in aj:
                pend(info, "VR", f"Voltou a ter direito (no mês anterior: "
                                 f"{b.observacao})",
                     "Recalculado. Confira.", False)
            if "quantidade" in aj:
                qtde = aj["quantidade"]
            elif zero_regra:
                qtde = 0
            elif carrega_zero:
                qtde = 0
                pend(info, "VR", f"Estava zerado no mês anterior"
                                 f"{': ' + b.observacao if b.observacao else ''}",
                     "Mantido zerado. Se voltou a ter direito, informe a "
                     "quantidade em ajustes.", False)
            else:
                qtde = base_dias
            fer, fal = ded_ferias(chave), ded_faltas(chave)
            af = dias_af if (sem_dir and not inteiro) else 0
            if af:
                obs.append(f"Afastado parte do mês: {txt_af}")
            deducoes = fer + fal + af
        if zero_regra:
            obs.insert(0, zero_regra)
        elif b is not None and b.qtde_fixa == 0 and qtde == 0 and b.observacao \
                and "quantidade" not in aj and tipo != "PLANTÃO":
            obs.insert(0, b.observacao)        # mantém o motivo de mês a mês
        if not sem_dir and txt_af:
            obs.append(f"Afastamento que não tira o direito: {txt_af}")
        if _contem(r.get("situacao"), regras["situacoes_verificar"]):
            pend(info, "VR", f"Situação: {info['Situação']}",
                 "Verifique se tem direito no mês.")
            travas.append("situacao")
        if b is not None and _eh_educacao(r, regras):
            pend(info, "VR", "Lotado na Educação no SIGA, mas estava na planilha "
                             "do mês anterior",
                 "Confira a lotação (o VR da Educação é feito pela SEMED).")
            travas.append("educacao")

        if b is not None and tipo != "PLANTÃO" and "quantidade" not in aj \
                and b.qtde_fixa not in (None, 0, qtde_comum, du):
            pend(info, "VR", f"No mês anterior a QTDE foi digitada à mão "
                             f"({b.qtde_fixa:g}{': ' + b.observacao if b.observacao else ''})",
                 f"Recalculado ({qtde:g}). Se o valor à mão ainda vale, informe "
                 f"em ajustes.", False)
        if b is not None and b.total_fixo is not None and "quantidade" not in aj:
            pend(info, "VR", f"No mês anterior o TOTAL foi digitado à mão "
                             f"({_br(b.total_fixo)})",
                 "Recalculado pela fórmula. Confira.", False)
        desc, reemb = aj.get("desconto", 0) or 0, aj.get("reembolso", 0) or 0
        if aj.get("observacao"):
            obs.append(aj["observacao"])
        total = None
        if valor_dia is not None:
            total = round(max(0, qtde - deducoes) * valor_dia - desc + reemb, 2)
        if _txt(r.get("cpf")) and (total or 0) > 0:
            cpf_vr.setdefault(_txt(r.get("cpf")), []).append(info)
        vr.append({**info, "Origem": origem, "Tipo": tipo, "Dias úteis": du,
                   "Afastamento": af, "Faltas": fal, "Férias (dias úteis)": fer,
                   "Total deduções": deducoes, "Quantidade": qtde,
                   "Valor diário": valor_dia, "Desconto": desc,
                   "Reembolso": reemb, "Valor total": total,
                   "Observação": "; ".join(obs), "Obs. mês anterior": obs_ant,
                   "_padrao": (tipo != "PLANTÃO" and "quantidade" not in aj
                               and not zero_regra and qtde == du
                               and base_dias == du),
                   "Status": "PENDENTE" if travas else "OK"})
        return tipo, valor_dia, bool(travas), bool(zero_regra)

    def complementar(info, adm, tipo, valor_dia, travado, zerado):
        """Complementar de quem entrou no mês M-2 (N2/N3)."""
        if zerado or adm is None or not (ini_c2 <= adm <= fim_c2):
            return
        linha = {**info, "Tipo": tipo, "Valor diário": valor_dia}
        if tipo == "PLANTÃO":
            pend(info, "Complementar", "Plantonista admitido: complementar "
                                       "não calculada",
                 "Calcule pelos plantões efetivamente feitos.")
            comp.append({**linha, "Status": "PENDENTE"})
            return
        d1, d2 = cal.dias_uteis(adm, fim_c2), du_c1
        comp.append({**linha, "Dias retroativos (mês da admissão)": d1,
                     "Dias do mês seguinte": d2, "Total de dias": d1 + d2,
                     "Valor total": round((d1 + d2) * valor_dia, 2)
                     if valor_dia is not None else None,
                     "Status": "PENDENTE" if travado else "OK"})

    # 1) quem estava na planilha do mês anterior
    for b in base:
        r = siga.get(b.chave)
        if r is None:
            info = {"Matrícula": b.matricula, "Nome": b.nome, "Cargo": "",
                    "Secretaria": "", "Vínculo": "", "Admissão": "",
                    "Situação": ""}
            excl.append({**info, "Benefício": "VR",
                         "Motivo": "Não está no relatório do SIGA "
                                   "(exonerado/desligado?)"})
            pend(info, "VR", "Estava na planilha do mês anterior, mas não está "
                             "no relatório do SIGA",
                 "Saiu da planilha. Confira se foi exonerado; se não, confira o "
                 "relatório do SIGA.", False)
            continue
        info = _info(r)
        adm = parse_data(r.get("data_admissao"))
        exon = parse_data(r.get("data_exoneracao"))
        motivo = None
        if _contem(r.get("situacao"), regras["situacoes_excluir"]):
            motivo = f"Situação: {info['Situação']}"
        elif exon and exon < ini_comp:
            motivo = f"Exonerado/fim de contrato em {exon:%d/%m/%Y}"
        else:
            motivo = _motivo_exclusao(r, adm, vr_cfg["exclusoes"])
        if motivo:
            excl.append({**info, "Benefício": "VR", "Motivo": motivo})
            pend(info, "VR", f"Saiu da planilha: {motivo}",
                 "Confira. Estava na planilha do mês anterior.", False)
            continue
        complementar(info, adm, *linha_vr(b.chave, r, b, "planilha anterior"))

    # 2) quem não estava: novos admitidos, inclusões manuais, demais ativos
    na_base = {b.chave for b in base}
    for chave, r in siga.items():
        if chave in na_base:
            continue
        info = _info(r)
        adm = parse_data(r.get("data_admissao"))
        exon = parse_data(r.get("data_exoneracao"))
        if _contem(r.get("situacao"), regras["situacoes_excluir"]) or \
                (exon and exon < ini_comp):
            continue
        incluir_manual = bool(ajustes_por.get(chave, {}).get("tipo"))
        educa = _eh_educacao(r, regras)
        if educa and not incluir_manual:                    # D1/D2
            if adm and adm >= ini_c2:
                edu.append({**info, "Ciclo": "novo admitido"})
            continue
        motivo = _motivo_exclusao(r, adm, vr_cfg["exclusoes"])
        recente = adm is not None and adm >= ini_c2
        if motivo:
            if recente:
                excl.append({**info, "Benefício": "VR", "Motivo": motivo})
            continue
        if adm is None and not incluir_manual:
            pend(info, "VR", "Sem data de admissão no SIGA",
                 "Confira o cadastro.", False)
            continue
        if adm and adm >= ini_c1 and not incluir_manual:     # N2
            aguard.append({**info, "Observação":
                           f"Admitido em {adm:%m/%Y}: complementar e planilha "
                           f"entram no próximo ciclo"})
            continue
        if not recente and not incluir_manual:
            fora.append({**info, "Tipo sugerido": sugerir_tipo(chave, r)[0],
                         "Salário": _txt(r.get("salario"))})
            continue
        origem = "novo admitido" if recente else "incluído por ajuste"
        resultado = linha_vr(chave, r, None, origem)
        if recente:
            complementar(info, adm, *resultado)

    for chave in ajustes_por:
        if chave not in siga and chave not in na_base:
            avisos.append(f"Ajuste da matrícula {chave} ignorado: ela não está "
                          f"no SIGA nem na planilha.")

    for lista in cpf_vr.values():                       # R10: duplo vínculo
        if len(lista) > 1:
            mats = ", ".join(i["Matrícula"] for i in lista)
            for i in lista:
                pend(i, "VR", f"Duplo vínculo (matrículas {mats})",
                     "Pelo POP, o VR considera a soma das jornadas. Confira o "
                     "valor de cada matrícula.", False)

    # ---------------------------------------------------------- VA (A1-A5)
    va = []
    if va_cfg.get("ativo"):
        credito = numero(va_cfg.get("valor_credito"))
        desconto = numero(va_cfg.get("valor_desconto")) or 0
        if credito is None:
            avisos.append("Valor do crédito do VA não configurado.")
        for chave, r in siga.items():
            info = _info(r)
            adm = parse_data(r.get("data_admissao"))
            exon = parse_data(r.get("data_exoneracao"))
            if _eh_educacao(r, regras) and not va_cfg.get("inclui_educacao"):
                continue
            if _contem(r.get("situacao"), regras["situacoes_excluir"]) or \
                    (exon and exon < ini_comp) or adm is None or adm >= ini_c1:
                continue
            motivo = _motivo_exclusao(r, adm, va_cfg["exclusoes"])
            if motivo:
                excl.append({**info, "Benefício": "VA", "Motivo": motivo})
                continue
            sem_dir, dias_af, inteiro, txt_af = afastamento(
                r, va_cfg.get("afastamentos_sem_direito") or [])
            if sem_dir and inteiro:
                excl.append({**info, "Benefício": "VA",
                             "Motivo": f"Afastado o mês todo: {txt_af}"})
                continue
            status, obs = "OK", []
            if sem_dir and dias_af:
                pend(info, "VA", f"Afastado parte do mês ({txt_af})",
                     "Confira se recebe o VA.", False)
                obs.append(f"Afastado parte do mês: {txt_af}")
            if _contem(r.get("situacao"), regras["situacoes_verificar"]):
                pend(info, "VA", f"Situação: {info['Situação']}",
                     "Verifique se tem direito no mês.")
                status = "PENDENTE"
            nf = int(faltas_por.get(chave, 0))
            cred, desc = credito, desconto
            if nf:
                regra = va_cfg.get("com_falta", "perde")
                if regra == "perde":
                    cred, desc = 0, 0
                    obs.append(f"{nf} falta(s) injustificada(s): perde o VA do "
                               f"mês (sem o desconto)")
                elif regra == "manter":
                    obs.append(f"{nf} falta(s): mantido")
                else:
                    pend(info, "VA", f"{nf} falta(s) no período",
                         "Decida se perde o VA (Configurações > Valores).")
                    status = "PENDENTE"
            va.append({**info, "Faltas": nf, "Crédito": cred, "Desconto": desc,
                       "Observação": "; ".join(obs), "Status": status})

    if comp:
        avisos.append(f"Complementar calculada para os admitidos em "
                      f"{m2:02d}/{a2}: confira as exceções antes de lançar (N3).")

    ordem = {"VR": 0, "VR/VA": 1, "Complementar": 2, "VA": 3}
    pend_df = pd.DataFrame(pendencias, columns=COLS_PEND)
    if len(pend_df):
        pend_df = pend_df.assign(_t=pend_df["Trava a importação"] != "Sim",
                                 _b=pend_df["Benefício"].map(ordem).fillna(9)) \
            .sort_values(["_t", "_b", "Nome"]).drop(columns=["_t", "_b"]) \
            .reset_index(drop=True)
    return Resultado(
        ano=ano, mes=mes, dias_uteis=du, folgas=cal.folgas_no_mes(ano, mes),
        vr=pd.DataFrame(vr, columns=COLS_VR + ["_padrao"]),
        va=pd.DataFrame(va, columns=COLS_VA),
        complementar=pd.DataFrame(comp, columns=COLS_COMP),
        excluidos=pd.DataFrame(excl, columns=COLS_INFO + ["Benefício", "Motivo"]),
        educacao=pd.DataFrame(edu, columns=COLS_INFO + ["Ciclo"]),
        aguardando=pd.DataFrame(aguard, columns=COLS_INFO + ["Observação"]),
        fora=pd.DataFrame(fora, columns=COLS_INFO + ["Tipo sugerido", "Salário"]),
        pendencias=pend_df, avisos=avisos, valores_tipo=valores,
        ferias_usadas=pd.DataFrame(ferias_linhas, columns=[
            "Matrícula", "Nome", "Período de gozo", "Início ajustado",
            "Fim ajustado", "Dias úteis"]),
        faltas_usadas=faltas_usadas, carga_anterior=str(arq_carga))


# ==========================================================================
# Saída
# ==========================================================================
def _gravar_excel(caminho, abas):
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        for nome, df in abas:
            nome = nome[:31]
            df = df.drop(columns=[c for c in df.columns if str(c).startswith("_")])
            df.to_excel(w, sheet_name=nome, index=False)
            ws = w.sheets[nome]
            ws.freeze_panes = "A2"
            for cel in ws[1]:
                cel.font = Font(bold=True)
            for i, col in enumerate(df.columns, 1):
                larg = max([len(str(col))] +
                           [len(_txt(v)) for v in df[col].head(300)])
                ws.column_dimensions[get_column_letter(i)].width = min(60, larg + 2)


def _layout(df, layout):
    """Monta o arquivo de importação com as colunas pedidas no config."""
    saida = pd.DataFrame()
    for titulo_col, campo in layout:
        rotulo = ROTULOS.get(campo)
        if rotulo is None or rotulo not in df:
            raise ErroEntrada(f"Configurações > Importação: campo '{campo}' "
                              f"não existe neste arquivo.")
        saida[titulo_col] = df[rotulo].values
    return saida


def nome_planilha_carga(ano, mes):
    """Nome no padrão da Divisão: '11- NOVEMBRO - 2026.xlsx' (mês da folha)."""
    af, mf = mes_antes(ano, mes, 1)
    return f"{mf:02d}- {MESES[mf - 1]} - {af}.xlsx"


def gravar_planilha_carga(res, caminho):
    """A planilha no formato da Divisão (aba CARGA com fórmulas, como a
    '11- NOVEMBRO - 2026'). Ela é a base do mês seguinte."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = "CARGA"
    negrito = Font(bold=True)
    af, mf = res.folha
    ws["A1"] = (f"CARGA VALE-REFEIÇÃO — folha de {MESES[mf - 1].lower()}/{af} "
                f"(uso em {res.mes:02d}/{res.ano}) — gerada pelo programa em "
                f"{pd.Timestamp.now():%d/%m/%Y %H:%M}")
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"], ws["B2"] = "INICIO", date(res.ano, res.mes, 1)
    ws["A3"], ws["B3"] = "FIM", ultimo_dia(res.ano, res.mes)
    ws["H3"], ws["I3"] = "TOTAL", res.dias_uteis
    ws["A4"] = "FERIADOS"
    ws["B4"] = "; ".join(f"{d:%d/%m} {desc}" for d, desc in res.folgas) or "nenhum"
    for c in ("B2", "B3"):
        ws[c].number_format = "DD/MM/YYYY"
    for c in ("A2", "A3", "A4", "H3"):
        ws[c].font = negrito
    tipos = list(res.valores_tipo.items())[:5]
    for i, (t, v) in enumerate(tipos):
        ws.cell(2 + i, 11, t).font = negrito
        ws.cell(2 + i, 12, v).number_format = '"R$" #,##0.00'
    faixa = f"$K$2:$L${1 + max(1, len(tipos))}"

    cab = ["MAT", "NOME", "SITUAÇÃO", "TIPO", "DEDUÇÃO AFASTAMENTO",
           "DEDUÇÃO DE FALTAS", "DEDUÇÃO DE FÉRIAS", "TOTAL DEDUÇÕES", "QTDE",
           "VALOR", "DESCONTO", "TOTAL", "REEMBOLSO", "OBSERVAÇÃO",
           "SECRETARIA", "CARGO", "ADMISSÃO", "STATUS"]
    lin_cab = 7
    for j, t in enumerate(cab, 1):
        ws.cell(lin_cab, j, t).font = negrito
    laranja = PatternFill("solid", fgColor="FFE0B2")
    moeda = '"R$" #,##0.00'
    vr = res.vr.reset_index(drop=True)
    for i, r in vr.iterrows():
        n = lin_cab + 1 + i
        mat = r["Matrícula"]
        ws.cell(n, 1, int(mat) if str(mat).isdigit() else mat)
        ws.cell(n, 2, r["Nome"])
        ws.cell(n, 3, r["Situação"])
        ws.cell(n, 4, r["Tipo"] or None)
        ws.cell(n, 5, r["Afastamento"] or 0)
        ws.cell(n, 6, r["Faltas"] or 0)
        ws.cell(n, 7, r["Férias (dias úteis)"] or 0)
        ws.cell(n, 8, f'=IF(D{n}="PLANTÃO",0,E{n}+F{n}+G{n})')
        ws.cell(n, 9, "=$I$3" if r["_padrao"] else r["Quantidade"])
        ws.cell(n, 10, f"=IFERROR(VLOOKUP(D{n},{faixa},2,FALSE),0)").number_format = moeda
        ws.cell(n, 11, r["Desconto"] or 0).number_format = moeda
        ws.cell(n, 12, f"=MAX(0,I{n}-H{n})*J{n}-K{n}+M{n}").number_format = moeda
        ws.cell(n, 13, r["Reembolso"] or 0).number_format = moeda
        ws.cell(n, 14, r["Observação"] or None)
        ws.cell(n, 15, r["Secretaria"])
        ws.cell(n, 16, r["Cargo"])
        ws.cell(n, 17, r["Admissão"])
        ws.cell(n, 18, r["Status"])
        if r["Status"] != "OK":
            for j in range(1, len(cab) + 1):
                ws.cell(n, j).fill = laranja
    ultima = lin_cab + max(1, len(vr))
    if len(vr):
        tab = Table(displayName="TabelaCarga", ref=f"A{lin_cab}:{get_column_letter(len(cab))}{ultima}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)
    ws["H2"], ws["I2"] = "SOMA", f"=SUM(L{lin_cab + 1}:L{ultima})"
    ws["H2"].font = negrito
    ws["I2"].number_format = moeda
    larguras = [9, 40, 16, 15, 12, 10, 10, 10, 7, 10, 11, 12, 11, 45, 35, 40, 11, 10]
    for j, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in range(1, len(cab) + 1):
        ws.cell(lin_cab, j).alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = f"C{lin_cab + 1}"

    def aba(nome, df):
        if df is None or not len(df):
            return
        w2 = wb.create_sheet(nome)
        df = df.drop(columns=[c for c in df.columns if str(c).startswith("_")])
        w2.append([str(c) for c in df.columns])
        for c in w2[1]:
            c.font = negrito
        for linha in df.itertuples(index=False):
            w2.append([None if (isinstance(v, float) and pd.isna(v)) else v
                       for v in linha])
        for j, col in enumerate(df.columns, 1):
            w2.column_dimensions[get_column_letter(j)].width = \
                min(45, max(10, len(str(col)) + 2))
        w2.freeze_panes = "A2"

    aba("FERIAS", res.ferias_usadas)
    aba("FALTAS", res.faltas_usadas)
    dados = wb.create_sheet("DADOS")
    for linha in [("Competência (uso)", f"{res.mes:02d}/{res.ano}"),
                  ("Folha (crédito)", f"{mf:02d}/{af}"),
                  ("Dias úteis", res.dias_uteis),
                  ("Base (mês anterior)", Path(res.carga_anterior).name)] + \
            [("Feriado/ponto facultativo", f"{d:%d/%m/%Y} {desc}")
             for d, desc in res.folgas]:
        dados.append(list(linha))
    dados.column_dimensions["A"].width = 26
    dados.column_dimensions["B"].width = 45
    wb.calculation.fullCalcOnLoad = True
    wb.save(caminho)


def salvar(res, pasta_saida, regras):
    """Grava as planilhas da competência. Devolve a lista de arquivos (a
    primeira é a de conferência)."""
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    r = res.rotulo
    vr = res.vr
    vr_ok = vr[(vr["Status"] == "OK") & (vr["Valor total"].fillna(0) > 0)]
    va = res.va
    va_ok = va[va["Status"] == "OK"]
    trava = (res.pendencias["Trava a importação"] == "Sim").sum()
    af, mf = res.folha

    resumo = [
        ("Competência (uso do benefício)", f"{res.mes:02d}/{res.ano}"),
        ("Folha (crédito no último dia útil)", f"{mf:02d}/{af}"),
        ("Gerado em", pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")),
        ("Base (planilha do mês anterior)", Path(res.carga_anterior).name),
        ("Dias úteis", res.dias_uteis),
        ("Feriados/pontos descontados",
         "; ".join(f"{d:%d/%m} {desc}" for d, desc in res.folgas) or "nenhum"),
        ("VR: servidores na planilha", len(vr)),
        ("VR: prontos para importar (valor > 0)", len(vr_ok)),
        ("VR: valor total (prontos)", round(float(vr_ok["Valor total"].sum()), 2)),
        ("VR: valor total (toda a planilha)",
         round(float(vr["Valor total"].fillna(0).sum()), 2)),
        ("VA: créditos prontos", int((va_ok["Crédito"].fillna(0) > 0).sum())),
        ("VA: valor dos créditos", round(float(va_ok["Crédito"].fillna(0).sum()), 2)),
        ("Complementar", len(res.complementar)),
        ("Pendências (travam a importação)", f"{len(res.pendencias)} ({trava})"),
        ("Excluídos", len(res.excluidos)),
        ("Aguardando próximo ciclo", len(res.aguardando)),
        ("Educação: novos admitidos a enviar", len(res.educacao)),
    ] + [("Aviso", a) for a in res.avisos] + [
        ("IMPORTANTE", "Confira as pendências antes de importar na Ábaco. "
                       "Linhas pendentes NÃO vão para os arquivos de importação. "
                       "Para resolver, preencha o arquivo de ajustes e calcule "
                       "de novo."),
    ]

    arquivos = []
    conf = pasta / f"VR_VA_{r}_conferencia.xlsx"
    _gravar_excel(conf, [
        ("Resumo", pd.DataFrame(resumo, columns=["Item", "Valor"])),
        ("Pendências", res.pendencias), ("VR", vr), ("VA", va),
        ("Complementar", res.complementar), ("Excluídos", res.excluidos),
        ("Aguardando próximo ciclo", res.aguardando),
        ("Educação", res.educacao), ("Fora da planilha", res.fora)])
    arquivos.append(conf)

    carga = pasta / nome_planilha_carga(res.ano, res.mes)
    gravar_planilha_carga(res, carga)
    arquivos.append(carga)

    imp = pasta / f"VR_{r}_importacao_abaco.xlsx"
    _gravar_excel(imp, [("VR", _layout(vr_ok, regras["vr"]["layout_importacao"]))])
    arquivos.append(imp)

    if regras["va"].get("ativo"):
        cred = va_ok[va_ok["Crédito"].fillna(0) > 0]
        desc = va_ok[va_ok["Desconto"].fillna(0) > 0]
        for nome, df, lay in (("credito", cred, "layout_credito"),
                              ("desconto", desc, "layout_desconto")):
            p = pasta / f"VA_{r}_{nome}_importacao_abaco.xlsx"
            _gravar_excel(p, [(f"VA {nome}", _layout(df, regras["va"][lay]))])
            arquivos.append(p)

    if len(res.complementar):
        cp = pasta / f"Complementar_{r}.xlsx"
        _gravar_excel(cp, [("Complementar", res.complementar)])
        arquivos.append(cp)

    if len(res.educacao):                                  # D2
        ed = pasta / f"Educacao_{r}_novos_admitidos.xlsx"
        _gravar_excel(ed, [("Novos admitidos", res.educacao[
            ["Matrícula", "Nome", "Admissão", "Cargo"]].rename(
                columns={"Admissão": "Data de admissão"}))])
        arquivos.append(ed)

    # modelo de ajustes com quem está pendente (preencher e salvar em entrada/)
    pend = res.pendencias
    if len(pend):
        tipos = dict(zip(vr["Matrícula"], vr["Tipo"]))
        g = pend.groupby("Matrícula", sort=False).agg(
            Nome=("Nome", "first"),
            Motivo=("Motivo", lambda m: " | ".join(dict.fromkeys(m))),
            Trava=("Trava a importação", lambda t: "Sim" if (t == "Sim").any() else "Não"))
        g = g.reset_index()
        g.insert(3, "Tipo atual", g["Matrícula"].map(tipos).fillna(""))
        for c in ("Tipo", "Quantidade", "Desconto", "Reembolso", "Observação"):
            g[c] = None
        g = g.rename(columns={"Trava": "Trava a importação"})
        aj = pasta / f"ajustes_para_preencher_{r}.xlsx"
        instr = pd.DataFrame({"Como usar": [
            "1. Na aba Ajustes, preencha só o que precisa mudar: Tipo (NORMAL, "
            "PLANTÃO, DIFERENCIADO, DIFERENCIADO 2), Quantidade (QTDE antes de "
            "férias/faltas; 0 = não faz jus), Desconto e Reembolso em R$.",
            "2. Apague as linhas que não precisam de ajuste (ou deixe em branco).",
            "3. Salve como 'ajustes.xlsx' na pasta de entrada (ou use 'Adicionar "
            "arquivo manualmente' > Ajustes) e calcule de novo.",
        ]})
        _gravar_excel(aj, [("Ajustes", g), ("Como usar", instr)])
        arquivos.append(aj)
    return arquivos


def gerar_modelos(pasta):
    """Cria planilhas vazias com as colunas esperadas."""
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    criados = []
    for nome, colunas in MODELOS.items():
        destino = pasta / f"{nome}.xlsx"
        _gravar_excel(destino, [(nome, pd.DataFrame(columns=colunas))])
        criados.append(destino)
    return criados
