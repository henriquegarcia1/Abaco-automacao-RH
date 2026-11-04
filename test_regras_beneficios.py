# -*- coding: utf-8 -*-
"""
Testes das regras de VR/VA com dados fictícios, no mesmo formato das
planilhas reais da Divisão de Benefícios:

  * carga_anterior.xlsx  aba CARGA como a '11- NOVEMBRO - 2026' (com fórmulas)
  * servidores.xlsx      relatório 'situação de afastamento' do SIGA
  * ferias.xlsx          relatório da Ábaco com 'Período de Gozo'
  * faltas.xlsx          relatório da Ábaco com 'Qtde'

Competência dos testes: 12/2026 (22 dias úteis), folha de 11/2026.

    python -m unittest test_regras_beneficios -v
"""

import tempfile
import unittest
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

import regras_beneficios as rb

SIGA_COLS = ["MATRICULA", "NOME", "CPF", "SITUACAO", "REGIME", "NOME_SECRETARIA",
             "CODIGO_SECRETARIA", "NOMEACAO", "EXONERACAO", "CARGO_EFETIVO",
             "CARGO_COMISSIONADO", "SALARIO", "SALARIO_INFORMADO",
             "INICIO_AFASTAMENTO", "TERMINO_AFASTAMENTO", "TIPO_AFASTAMENTO"]


def _serial(d):
    """Data no formato do SIGA (número de série do Excel)."""
    return float((d - date(1899, 12, 30)).days)


def siga(mat, nome, regime="CARREIRA", secretaria="SECRETARIA MUNICIPAL DE SAÚDE",
         adm=date(2015, 3, 2), exon=None, cargo="2635 - ASSISTENTE ADMINISTRATIVO - 40 HORAS",
         comissionado=" - ", salario="3045,76", situacao="EM EXERCÍCIO",
         afast=("", "", ""), cpf=None):
    return dict(zip(SIGA_COLS, [
        str(mat), nome, cpf or f"{int(mat):011d}", situacao, regime, secretaria,
        "12", _serial(adm), _serial(exon) if exon else "", cargo, comissionado,
        salario, "", afast[0], afast[1], afast[2]]))


def carga_anterior(caminho, linhas):
    """Imita a aba CARGA: cabeçalho na linha 7, QTDE '=$I$3' (padrão) ou
    número (digitado à mão)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "CARGA"
    ws["A2"], ws["B2"] = "INICIO", "2026-11-01"
    ws["H3"], ws["I3"] = "TOTAL", 19
    for i, (t, v) in enumerate([("NORMAL", 56), ("PLANTÃO", 84),
                                ("DIFERENCIADO", 42), ("DIFERENCIADO 2", 28)]):
        ws.cell(2 + i, 11, t)
        ws.cell(2 + i, 12, v)
    cab = ["MAT", "NOME", "SITUAÇÃO", "TIPO", "DEDUÇÃO ATESTATO",
           "DEDUÇÃO DE FALTAS", "DEDUÇÃO DE FÉRIAS", "TOTAL DEDUÇÕES", "QTDE",
           "VALOR", "DESCONTO", "TOTAL", "REEMBOLSO", "OBSERVAÇÃO"]
    for j, c in enumerate(cab, 1):
        ws.cell(7, j, c)
    for i, (mat, nome, tipo, qtde, obs) in enumerate(linhas):
        n = 8 + i
        ws.cell(n, 1, mat)
        ws.cell(n, 2, nome)
        ws.cell(n, 4, tipo)
        ws.cell(n, 9, "=$I$3" if qtde is None else qtde)
        ws.cell(n, 12, f"=I{n}*J{n}")
        ws.cell(n, 14, obs)
    wb.save(caminho)


def _xlsx(pasta, nome, linhas, titulo=None):
    df = pd.DataFrame(linhas)
    with pd.ExcelWriter(pasta / f"{nome}.xlsx", engine="openpyxl") as w:
        if titulo:   # relatório com título antes do cabeçalho
            pd.DataFrame([[titulo], [""]]).to_excel(w, index=False, header=False)
            df.to_excel(w, index=False, startrow=2)
        else:
            df.to_excel(w, index=False)


class TestCalendario(unittest.TestCase):
    def setUp(self):
        self.cal = rb.Calendario.do_config(2025, 2027, rb.REGRAS_PADRAO)

    def test_novembro_2026_tem_19_dias_uteis(self):          # R1 (reunião)
        self.assertEqual(self.cal.dias_uteis_mes(2026, 11), 19)

    def test_dezembro_2026_tem_22_dias_uteis(self):          # R1 (reunião)
        self.assertEqual(self.cal.dias_uteis_mes(2026, 12), 22)

    def test_ponto_facultativo_do_config_e_descontado(self):
        regras = rb.regras_completas({"pontos_facultativos":
                                      {"2026-12-24": "Véspera de Natal"}})
        cal = rb.Calendario.do_config(2026, 2026, regras)
        self.assertEqual(cal.dias_uteis_mes(2026, 12), 21)

    def test_pascoa(self):
        self.assertEqual(rb.pascoa(2026), date(2026, 4, 5))
        self.assertEqual(rb.pascoa(2027), date(2027, 3, 28))


class TestNormalizacao(unittest.TestCase):
    def test_jornada(self):
        for txt, esperado in [("4h30", "4h30"), ("4:30", "4h30"),
                              ("4,5 horas", "4h30"), ("6 horas", "6h"),
                              ("06:00", "6h"), ("12x36", "12h"),
                              ("40 horas semanais", "40h"), ("40.0", "40h")]:
            self.assertEqual(rb.chave_jornada(txt), esperado, txt)

    def test_horas_no_nome_do_cargo(self):
        self.assertEqual(rb.horas_do_texto("2631 - VIGIA MUNICIPAL - 40 HORAS"), "40")
        self.assertEqual(rb.horas_do_texto("2605 - ASSISTENTE SOCIAL-30HS"), "30")
        self.assertEqual(rb.horas_do_texto("2448 - ODONTOLOGO GENERALISTA - 20 HS"), "20")
        self.assertEqual(rb.horas_do_texto("1028 - ASSESSOR II"), "")

    def test_matricula(self):
        self.assertEqual(rb.chave_matricula("017446"), "17446")
        self.assertEqual(rb.chave_matricula("17446.0"), "17446")

    def test_salario_do_siga(self):
        self.assertEqual(rb.numero("5498,14"), 5498.14)
        self.assertEqual(rb.numero("5.498,14"), 5498.14)
        self.assertEqual(rb.numero("2534.14"), 2534.14)

    def test_periodo_de_gozo(self):
        self.assertEqual(rb.datas_no_texto("23/11/2026 à 02/12/2026"),
                         [date(2026, 11, 23), date(2026, 12, 2)])

    def test_config_antigo_e_migrado(self):
        antigo = {"feriados_municipais": {"2026-06-04": "X"},
                  "vr": {"valor_dia_por_jornada": {"40h": None}},
                  "exclusoes": [{"campo": "vinculo", "contem": "agente politico"}]}
        r = rb.regras_completas(antigo)
        self.assertEqual(r["versao"], rb.VERSAO_REGRAS)
        self.assertEqual(r["feriados_municipais"], {"2026-06-04": "X"})
        self.assertNotIn("valor_dia_por_jornada", r["vr"])
        self.assertEqual(r["vr"]["valor_dia_por_tipo"]["NORMAL"], 56.0)


class TestCalculo(unittest.TestCase):
    """Competência 12/2026: 22 dias úteis; admitidos em 10/2026 entram agora
    com complementar; admitidos em 11/2026 aguardam o próximo ciclo."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)
        carga_anterior(self.pasta / "carga_anterior.xlsx", [
            (1001, "ANA ADM", "NORMAL", None, None),
            (1002, "BRUNO FERIAS", "NORMAL", None, None),
            (1003, "CARLA FALTAS", "NORMAL", 19, None),            # 19 digitado (nov)
            (1004, "DIEGO PLANTAO", "PLANTÃO", 15, None),
            (1005, "ELISA PLANTAO SEM PLANILHA", "PLANTÃO", 14, None),
            (1006, "FABIO DIF RICO", "DIFERENCIADO", None, None),
            (1007, "GABI DIF2", "DIFERENCIADO 2", None, None),
            (1008, "HUGO TRT", "NORMAL", 0, "OPTOU POR RECEBIMENTO TRT"),
            (1009, "IVO INSS", "NORMAL", None, None),
            (1010, "JOAO EXONERADO", "NORMAL", None, None),
            (1011, "KATIA CONTRATO ACABA", "NORMAL", None, None),
            (1012, "LUIS AGENTE POLITICO", "NORMAL", None, None),
            (1013, "MARA MATERNIDADE", "NORMAL", None, None),
            (1014, "NICE OUTUBRO JA NA PLANILHA", "NORMAL", None, None),
        ])
        servidores = [
            siga(1001, "ANA ADM"),
            siga(1002, "BRUNO FERIAS"),
            siga(1003, "CARLA FALTAS"),
            siga(1004, "DIEGO PLANTAO", cargo="2631 - VIGIA MUNICIPAL - 40 HORAS"),
            siga(1005, "ELISA PLANTAO SEM PLANILHA",
                 cargo="2631 - VIGIA MUNICIPAL - 40 HORAS"),
            siga(1006, "FABIO DIF RICO", cargo="2585 - ENGENHEIRO-30HS",
                 salario="7784,43"),
            siga(1007, "GABI DIF2", cargo="2611 - PSICÓLOGO - 20 HORAS",
                 salario="4500,00"),
            siga(1008, "HUGO TRT"),
            siga(1009, "IVO INSS", situacao="AFASTADO",
                 afast=("30/10/2026", "", "LICENÇA PREVIDÊNCIA")),
            # 1010 não está mais no SIGA (exonerado)
            siga(1011, "KATIA CONTRATO ACABA", regime="TEMPORARIO",
                 exon=date(2026, 12, 18)),
            siga(1012, "LUIS AGENTE POLITICO", regime="AGENTES POLITICOS",
                 secretaria="GABINETE DO PREFEITO"),
            siga(1013, "MARA MATERNIDADE", situacao="AFASTADO",
                 afast=("01/11/2026", "28/02/2027", "LICENÇA MATERNIDADE")),
            siga(1014, "NICE OUTUBRO JA NA PLANILHA", adm=date(2026, 10, 5)),
            # novos admitidos em 10/2026 (complementar)
            siga(2001, "NINA COMISSIONADA", regime="COMISSIONADO", cargo=" - ",
                 comissionado="1028 - ASSESSOR II", adm=date(2026, 10, 26)),
            siga(2002, "OTAVIO VIGIA NOVO", regime="TEMPORARIO",
                 cargo="2631 - VIGIA MUNICIPAL - 40 HORAS", adm=date(2026, 10, 15)),
            siga(2003, "PAULA PSICOLOGA NOVA", regime="TEMPORARIO",
                 cargo="2947 - PSICÓLOGO - TEMPORÁRIO - 30 HORAS",
                 adm=date(2026, 10, 19)),
            siga(2004, "QUEILA ESTAGIARIA", regime="ESTAGIARIO", cargo=" - ",
                 adm=date(2026, 10, 1)),
            # admitido em 11/2026: próximo ciclo
            siga(2005, "RAFAEL NOVEMBRO", adm=date(2026, 11, 10)),
            # Educação: fora do VR, mas recebe VA; novo admitido vai para a lista
            siga(3001, "SARA PROFESSORA", secretaria="SECRETARIA MUNICIPAL DE EDUCACAO",
                 cargo="2280 - PROFESSOR DA EDUCAÇÃO BÁSICA - PEB"),
            siga(3002, "TIAGO PROF NOVO", secretaria="SECRETARIA MUNICIPAL DE EDUCACAO",
                 cargo="2280 - PROFESSOR DA EDUCAÇÃO BÁSICA - PEB",
                 adm=date(2026, 10, 5)),
            # PSS de 2022: recebe VR? não está na planilha; VA não
            siga(3003, "URSULA PSS 2022", regime="PROC SEL PUBLICO",
                 adm=date(2022, 5, 2), secretaria="SECRETARIA MUNICIPAL DE EDUCACAO"),
        ]
        _xlsx(self.pasta, "servidores", servidores)
        _xlsx(self.pasta, "ferias", [
            {"Matricula ": "1002", "Nome do Servidor ": "BRUNO FERIAS",
             "Tipo de Folha": "NORMAL", "Período de Gozo": "23/11/2026 à 02/12/2026"},
        ], titulo="RELATÓRIO DE PERÍODO DE FÉRIAS")
        _xlsx(self.pasta, "faltas", [
            {"Nome": "CARLA FALTAS", "Matricula": "1003", "Cnt.": 1,
             "Competência": "01/10/2026", "Qtde": 3.6},
            {"Nome": "DIEGO PLANTAO", "Matricula": "1004", "Cnt.": 1,
             "Competência": "01/10/2026", "Qtde": 2},
            {"Nome": "SARA PROFESSORA", "Matricula": "3001", "Cnt.": 1,
             "Competência": "01/10/2026", "Qtde": 1},
        ])
        _xlsx(self.pasta, "plantoes", [
            {"Matrícula": "1004", "Nome": "DIEGO PLANTAO", "Quantidade de plantões": 16},
        ])
        self.regras = rb.regras_completas({})

    def tearDown(self):
        self.tmp.cleanup()

    def calcular(self):
        return rb.calcular(self.pasta, 2026, 12, self.regras)

    @staticmethod
    def linha(df, mat):
        sel = df[df["Matrícula"] == str(mat)]
        return None if sel.empty else sel.iloc[0]

    def pendencias(self, res, mat):
        p = res.pendencias
        return p[p["Matrícula"] == str(mat)]

    # ---- VR
    def test_normal_recebe_dias_uteis_vezes_valor(self):        # R1/R7
        r = self.linha(self.calcular().vr, 1001)
        self.assertEqual((r["Quantidade"], r["Valor total"], r["Status"]),
                         (22, 22 * 56.0, "OK"))

    def test_ferias_desconta_so_dias_uteis_da_competencia(self):  # R2/R4
        r = self.linha(self.calcular().vr, 1002)
        self.assertEqual(r["Férias (dias úteis)"], 2)          # 01 e 02/12
        self.assertEqual(r["Valor total"], 20 * 56.0)

    def test_falta_fracionada_e_truncada_como_na_planilha(self):  # R2
        r = self.linha(self.calcular().vr, 1003)
        self.assertEqual(r["Faltas"], 3)                        # INT(3,6)
        self.assertEqual(r["Quantidade"], 22)                   # 19 era de novembro
        self.assertEqual(r["Valor total"], 19 * 56.0)

    def test_plantonista_usa_planilha_de_plantoes_sem_deducao(self):  # R6
        r = self.linha(self.calcular().vr, 1004)
        self.assertEqual((r["Quantidade"], r["Total deduções"], r["Valor total"]),
                         (16, 0, 16 * 84.0))

    def test_plantonista_sem_planilha_mantem_o_mes_anterior(self):     # R6
        res = self.calcular()
        r = self.linha(res.vr, 1005)
        self.assertEqual((r["Quantidade"], r["Status"]), (14, "OK"))
        p = self.pendencias(res, 1005)
        self.assertTrue((p["Trava a importação"] == "Não").all() and len(p))

    def test_diferenciado_acima_do_limite_nao_faz_jus(self):   # R8
        res = self.calcular()
        r = self.linha(res.vr, 1006)
        self.assertEqual((r["Quantidade"], r["Valor total"]), (0, 0))
        self.assertIn("NÃO FAZ JUS", r["Observação"])
        r = self.linha(res.vr, 1007)
        self.assertEqual(r["Valor total"], 22 * 28.0)

    def test_zerado_a_mao_continua_zerado(self):                # obs. da planilha
        res = self.calcular()
        r = self.linha(res.vr, 1008)
        self.assertEqual(r["Valor total"], 0)
        self.assertEqual(r["Observação"], "OPTOU POR RECEBIMENTO TRT")

    def test_afastado_pelo_inss_o_mes_todo_zera(self):          # E4
        r = self.linha(self.calcular().vr, 1009)
        self.assertEqual(r["Valor total"], 0)
        self.assertTrue(r["Observação"].startswith("AFASTADO: LICENÇA PREVIDÊNCIA"))

    def test_licenca_maternidade_mantem(self):                  # E4
        r = self.linha(self.calcular().vr, 1013)
        self.assertEqual(r["Valor total"], 22 * 56.0)

    def test_fora_do_siga_sai_da_planilha(self):                # E6
        res = self.calcular()
        self.assertIsNone(self.linha(res.vr, 1010))
        self.assertIn("SIGA", self.linha(res.excluidos, 1010)["Motivo"])

    def test_contrato_que_termina_no_mes_e_proporcional(self):
        r = self.linha(self.calcular().vr, 1011)
        self.assertEqual(r["Quantidade"], 14)                   # 01 a 18/12

    def test_agente_politico_recebe_vr_mas_nao_va(self):        # E2 (POP)
        res = self.calcular()
        self.assertEqual(self.linha(res.vr, 1012)["Valor total"], 22 * 56.0)
        self.assertIsNone(self.linha(res.va, 1012))

    # ---- novos admitidos e complementar
    def test_comissionado_novo_entra_como_normal_com_complementar(self):  # N2
        res = self.calcular()
        r = self.linha(res.vr, 2001)
        self.assertEqual((r["Tipo"], r["Status"], r["Origem"]),
                         ("NORMAL", "OK", "novo admitido"))
        c = self.linha(res.complementar, 2001)
        # 26 a 30/10 = 5 dias úteis + novembro inteiro = 19
        self.assertEqual((c["Dias retroativos (mês da admissão)"],
                          c["Dias do mês seguinte"], c["Valor total"]),
                         (5, 19, 24 * 56.0))

    def test_admitido_em_m2_ja_na_planilha_tambem_tem_complementar(self):  # N2
        c = self.linha(self.calcular().complementar, 1014)
        # 05 a 30/10 = 19 dias úteis (12/10 é feriado) + novembro = 19
        self.assertEqual((c["Total de dias"], c["Status"]), (38, "OK"))

    def test_novo_admitido_aprende_o_tipo_pelo_cargo(self):     # N6
        r = self.linha(self.calcular().vr, 2002)
        # só 2 vigias na base: pouco para aprender -> sugestão pela jornada, trava
        self.assertEqual((r["Tipo"], r["Status"]), ("NORMAL", "PENDENTE"))

    def test_novo_admitido_sem_certeza_trava_ate_o_ajuste(self):
        res = self.calcular()
        self.assertEqual(self.linha(res.vr, 2003)["Status"], "PENDENTE")
        self.assertEqual(self.linha(res.complementar, 2003)["Status"], "PENDENTE")
        _xlsx(self.pasta, "ajustes", [{"Matrícula": "2003", "Tipo": "DIFERENCIADO"}])
        res = self.calcular()
        r = self.linha(res.vr, 2003)
        self.assertEqual((r["Tipo"], r["Status"], r["Valor total"]),
                         ("DIFERENCIADO", "OK", 22 * 42.0))

    def test_estagiario_e_novembro_e_educacao(self):            # E1/N2/D1
        res = self.calcular()
        self.assertIsNone(self.linha(res.vr, 2004))
        self.assertIsNotNone(self.linha(res.aguardando, 2005))
        self.assertIsNone(self.linha(res.vr, 3001))
        self.assertIsNotNone(self.linha(res.educacao, 3002))

    def test_ajustes_desconto_reembolso_e_quantidade(self):
        _xlsx(self.pasta, "ajustes", [
            {"Matrícula": "1001", "Desconto": "10,00", "Reembolso": "56,00"},
            {"Matrícula": "1008", "Quantidade": 22, "Observação": "voltou"},
        ])
        res = self.calcular()
        self.assertEqual(self.linha(res.vr, 1001)["Valor total"], 22 * 56.0 - 10 + 56)
        self.assertEqual(self.linha(res.vr, 1008)["Valor total"], 22 * 56.0)

    # ---- VA
    def test_va_credito_e_desconto(self):                       # A1/A3
        r = self.linha(self.calcular().va, 1001)
        self.assertEqual((r["Crédito"], r["Desconto"]), (400.0, 1.0))

    def test_va_falta_perde_tudo_sem_desconto(self):            # A2
        res = self.calcular()
        r = self.linha(res.va, 1003)
        self.assertEqual((r["Crédito"], r["Desconto"]), (0, 0))
        r = self.linha(res.va, 3001)                            # Educação também
        self.assertEqual(r["Crédito"], 0)

    def test_va_exclusoes(self):                                # E2/E3
        res = self.calcular()
        self.assertIsNone(self.linha(res.va, 3003))             # PSS 2022
        self.assertIsNone(self.linha(res.va, 2005))             # admitido em 11

    def test_va_continua_no_inss_e_sai_na_licenca_sem_vencimento(self):  # A6
        r = self.linha(self.calcular().va, 1009)                # Lei 3.035/2023
        self.assertEqual(r["Crédito"], 400.0)
        siga_extra = siga(1015, "VERA SEM VENCIMENTO", situacao="AFASTADO",
                          afast=("01/01/2026", "", "LICENÇA SEM VENCIMENTO"))
        df = pd.read_excel(self.pasta / "servidores.xlsx", dtype=str)
        pd.concat([df, pd.DataFrame([siga_extra]).astype(str)]).to_excel(
            self.pasta / "servidores.xlsx", index=False)
        self.assertIsNone(self.linha(self.calcular().va, 1015))

    def test_cedido_com_onus_do_municipio_recebe_va(self):    # Lei 3.035/2023
        siga_extra = siga(1016, "WALTER CEDIDO", regime="CEDIDO REMUN INTEG MUNIICIPIO")
        df = pd.read_excel(self.pasta / "servidores.xlsx", dtype=str)
        pd.concat([df, pd.DataFrame([siga_extra]).astype(str)]).to_excel(
            self.pasta / "servidores.xlsx", index=False)
        self.assertEqual(self.linha(self.calcular().va, 1016)["Crédito"], 400.0)

    def test_config_versao_2_remove_exclusao_de_cedido(self):
        v2 = {"versao": 2, "va": {"exclusoes": [
            {"campo": "vinculo", "contem": "cedid", "motivo": "x"},
            {"campo": "vinculo", "contem": "politic", "motivo": "y"}]}}
        r = rb.regras_completas(v2)
        self.assertEqual([e["contem"] for e in r["va"]["exclusoes"]], ["politic"])
        self.assertIn("sem vencimento", r["va"]["afastamentos_sem_direito"])

    # ---- saída
    def test_planilha_gerada_vira_a_base_do_mes_seguinte(self):
        res = self.calcular()
        saida = self.pasta / "saida"
        arquivos = rb.salvar(res, saida, self.regras)
        nomes = [a.name for a in arquivos]
        self.assertIn("11- NOVEMBRO - 2026.xlsx", nomes)
        self.assertIn("VA_2026-12_desconto_importacao_abaco.xlsx", nomes)

        carga = saida / "11- NOVEMBRO - 2026.xlsx"
        ws = load_workbook(carga)["CARGA"]
        self.assertEqual(ws["I3"].value, 22)
        cab = [c.value for c in ws[7]]
        self.assertEqual(cab[:4], ["MAT", "NOME", "SITUAÇÃO", "TIPO"])
        # QTDE padrão fica como fórmula; plantão como número
        qt = {str(ws.cell(n, 1).value): ws.cell(n, 9).value
              for n in range(8, ws.max_row + 1)}
        self.assertEqual(qt["1001"], "=$I$3")
        self.assertEqual(qt["1004"], 16)

        base, _ = rb.ler_carga_anterior(carga)
        tipos = {b.matricula: b.tipo for b in base}
        self.assertEqual(tipos["2001"], "NORMAL")
        self.assertEqual(tipos["1004"], "PLANTÃO")

        imp = pd.read_excel(saida / "VR_2026-12_importacao_abaco.xlsx", dtype=str)
        self.assertEqual(list(imp.columns), ["Matrícula", "Valor"])
        self.assertNotIn("1006", set(imp["Matrícula"]))         # zerado
        self.assertNotIn("2003", set(imp["Matrícula"]))         # pendente

    def test_sem_planilha_anterior_explica(self):
        (self.pasta / "carga_anterior.xlsx").unlink()
        with self.assertRaises(rb.ErroEntrada) as ctx:
            self.calcular()
        self.assertIn("MÊS ANTERIOR", str(ctx.exception))

    def test_novos_admitidos_para_conferir(self):
        novos = dict(rb.novos_admitidos(self.pasta, 2026, 12, self.regras))
        self.assertIn("2001", novos)
        self.assertNotIn("1001", novos)
        self.assertNotIn("3002", novos)                         # Educação


if __name__ == "__main__":
    unittest.main()
