# Regras de negócio — Benefícios (VR, VA e VT)

Fontes, nesta ordem de força:

1. **Reunião** com a Divisão de Benefícios em 03/11/2026 (transcrição:
   `Transcricao - Voz 261103_091622.txt`; horários entre colchetes).
2. **POP e fluxo** de VA e VR (`Cópia de Processo Vale Alimentação e
   Refeição e pop- Revisado 23_07.xlsx`, revisão de julho/2024).
3. **Planilha de carga** que a Divisão usa (`11- NOVEMBRO - 2026.xlsx`):
   fórmulas e observações das abas CARGA, FERIAS, FALTAS e DADOS.
4. **Relatório do SIGA** `Relatorio admissão situação de afastamento`
   (5.538 servidores), cruzado com a planilha.
5. **Legislação municipal** (pesquisa de 04/11/2026; veja a seção 12):
   **Lei 3.035/2023**, que alterou os arts. 78 e 79-A do Estatuto do
   Servidor (Lei 2.590/2017) e é a base atual do VR e do VA, e **Lei
   2.729/2019**.

Legenda:
- ✅ **confirmado**: dito na reunião, escrito no POP ou visível nas
  fórmulas da planilha.
- ⚠️ **a validar**: interpretação a partir dos dados. Já está no programa
  e pode ser ajustada em **Configurações**. Confirme com a área.

---

## 1. Como a Divisão trabalha hoje (e o programa reproduz)

```
 planilha do mês anterior ──┐   (aba CARGA: MAT, NOME, TIPO, QTDE, OBSERVAÇÃO)
 SIGA: situação de afastamento ┤   quem continua, afastamentos, admitidos, salário
 Ábaco: férias (3 meses)      ┤
 Ábaco: faltas                ┤
 secretaria: plantões         ┘
            │
            ▼
 aba CARGA atualizada: QTDE = dias úteis (plantonista: plantões)
                       TOTAL = (QTDE − férias − faltas) × valor do tipo
                               − desconto + reembolso
            │
            ▼
 importação na Ábaco (matrícula + valor) até o dia 20
```

- A planilha **não é refeita** todo mês: é a do mês anterior,
  atualizada. Isso foi visto na planilha [06:23], que tem 2.837 servidores
  na aba CARGA, cada um com o seu TIPO. ✅
- **Competência × folha**: a planilha `11- NOVEMBRO - 2026` (folha de
  novembro) calcula o uso de **dezembro** (aba CARGA: 01/12 a 31/12, 22
  dias úteis). O crédito sai no último dia útil do mês da folha. No
  programa, a **competência** é o mês de uso, e a planilha gerada leva o
  nome do mês da folha (`11- NOVEMBRO - 2026.xlsx` para a competência
  12/2026). ✅

## 2. Calendário do mês

| # | Regra | Fonte |
|---|---|---|
| C1 | Documentação recebida **até o dia 5** [06:23]; contratos dos recém-admitidos até o dia 10 (POP VR). | ✅ |
| C2 | Relatórios de férias e faltas **até o dia 15**; importação na folha **até o dia 20** (POP). | ✅ |
| C3 | O benefício é **pago antecipado** [01:21]: a folha do mês M credita o uso do mês M+1. | ✅ |

## 3. Vale-refeição (VR)

| # | Regra | Fonte |
|---|---|---|
| R1 | QTDE = **dias úteis** da competência, sem feriados e pontos facultativos. Nov/2026 = **19**, dez/2026 = **22** [01:21, 06:11]. | ✅ |
| R2 | TOTAL = (QTDE − faltas − férias) × VALOR − DESCONTO + REEMBOLSO; se as deduções passam da QTDE, o total é 0 (fórmula da aba CARGA). | ✅ planilha |
| R3 | Férias: relatório de **3 meses** [04:25]; cada período (`Período de Gozo`, ex.: "23/11/2026 à 02/12/2026") é cortado no mês da competência e conta os dias úteis (aba FERIAS: INICIO/FIM AJUSTADO). | ✅ |
| R4 | O programa não conta feriados como dia de férias (a fórmula NETWORKDAYS da planilha conta). Só muda algo quando há feriado no meio das férias. | ⚠️ |
| R5 | Faltas: coluna **Qtde** do relatório da Ábaco, truncada e limitada aos dias úteis: `INT(MIN(Qtde; dias úteis))` (aba FALTAS). Usa as faltas **do mês anterior à folha** (na planilha de out/2026 eram as de 09/2026). Se a mesma matrícula aparece em mais de uma linha, o programa **soma** (a planilha usava só a primeira). | ✅ / ⚠️ soma |
| R6 | **Plantonista** (TIPO PLANTÃO): QTDE = plantões informados pela secretaria (geralmente 14 a 16), **sem dedução** de férias e faltas (`TOTAL DEDUÇÕES = 0` para PLANTÃO). A Lei 3.035/2023 (art. 78, III) diz que a escala 12x36 recebe "apenas nos dias de trabalho efetivo". Sem a planilha de plantões, mantém a quantidade do mês anterior e avisa. | ✅ |
| R7 | **Valor por dia de cada TIPO** (aba CARGA, K2:L5): **NORMAL R$ 56** (40h), **PLANTÃO R$ 84** (12h), **DIFERENCIADO R$ 42** (6h a 8h), **DIFERENCIADO 2 R$ 28** (menos de 6h) [06:23]. O valor é fixado por **decreto** (Lei 3.035/2023, art. 78, § 4º). As proporções batem com o acordo de 2023 (8h = 1; 12x36 = 1,5; menos de 8h = 0,75), quando o valor era R$ 22,50 por dia; o decreto com o valor atual não foi achado na internet. | ✅ planilha |
| R8 | Jornada **inferior a 8 horas** (DIFERENCIADO e DIFERENCIADO 2) só recebe VR com **vencimento até o limite fixado em decreto** (Lei 3.035/2023, art. 78, I). Jornada de 8h/40h e escala 12x36 recebem **independentemente do vencimento** (art. 78, II e III). O limite usado é **R$ 5.662,23**: R$ 5.250,00 (POP de 2024) corrigido pelas revisões gerais (R$ 5.434,00 em 2025 e × 1,042 após o reajuste de 4,2% aprovado em julho/2026). A conta bate com os valores da planilha, mas o decreto com o valor de 2026 não foi achado na internet. O programa compara com o **vencimento básico** (coluna SALARIO do SIGA), como diz a lei. | ✅ regra / ⚠️ valor |
| R9 | O arquivo importado na Ábaco tem **matrícula e valor** (POP); só entram as linhas sem pendência e com valor > 0. Matrícula repetida na importação faz a folha ficar com o menor valor (POP), por isso o programa nunca repete matrícula. | ✅ |
| R10 | **Duplo vínculo**: "considerar a soma das jornadas" (POP). Como cada matrícula é uma linha, o programa marca para conferência quando o mesmo CPF tem duas matrículas recebendo. | ⚠️ |
| R11 | Jornada extra (SEMAD, SEMOS, SEMSA...) e ressarcimento de falta abonada entram como **DESCONTO/REEMBOLSO** em R$ (POP; colunas da aba CARGA). No programa: arquivo de **ajustes**. | ✅ |

## 4. Vale-alimentação (VA)

| # | Regra | Fonte |
|---|---|---|
| A1 | **Crédito fixo** de R$ 400,00 por matrícula (POP: "atualmente R$ 400,00"; acordo com o SINDSERP de 21/08/2023, convalidado pela Lei 3.035/2023). O valor é fixado por decreto (art. 79-A, § 5º). Duas matrículas = duas recargas. | ✅ |
| A2 | **Uma falta injustificada já faz perder o VA do mês inteiro** (POP; Lei 3.035/2023, art. 79-A, § 3º: não é devido em "ausência injustificada"). O faltoso também **não tem o desconto**. Falta **justificada** não tira o VA (§ 2º). | ✅ |
| A3 | **Desconto de R$ 1,00** para quem recebe o crédito (POP; Lei 2.728/2019, art. 3º, mantido pelo acordo de 2023). O programa gera dois arquivos de importação: crédito e desconto. | ✅ |
| A4 | Ressarcimento de falta abonada: o servidor recebe dois valores no mês seguinte (POP). No programa: lançar à mão por enquanto. | ⚠️ |
| A5 | O VA vale para **todos os servidores, inclusive a Educação**: "a todos os servidores ... independentemente da jornada" (Lei 3.035/2023, art. 79-A). A separação da Educação no POP é só do VR. | ✅ |
| A6 | **Afastamentos e o VA**: o VA **continua** em férias, recesso escolar, ausência justificada, compensação de horas, **afastamento pelo INSS** e licença remunerada por doença de pessoa da família (art. 79-A, § 2º). **Não é devido** nas **licenças sem remuneração** (§ 3º). O programa tira o VA de quem está o mês todo em licença sem vencimento, vacância, pagamento suspenso ou invalidez. | ✅ |

## 5. Quem tem direito

| # | Regra | VR | VA | Fonte |
|---|---|---|---|---|
| E1 | Estagiários, aposentados, pensionistas, inativos e sentença judicial não recebem. | não | não | ✅ POP |
| E2 | **Agentes políticos**: recebem **VR** (POP VR e planilha, com 23 na aba CARGA), mas **não VA** (POP VA e reunião [08:46]). **Cedidos**: recebem VR e VA **quando o ônus da remuneração é do Município** (Lei 3.035/2023, arts. 78, § 3º, e 79-A, § 4º), que é o caso dos 5 "CEDIDO REMUN INTEG MUNICIPIO" do SIGA. Na reunião foi dito que cedidos não recebem: provavelmente se referia aos cedidos com ônus da origem. | sim | sim (ônus do Município) | ✅ lei |
| E3 | **PSS de 2022** não faz jus ao **VA** [08:46]; no VR há 18 deles na planilha, então continuam. Critério: REGIME "PROC SEL PUBLICO" com nomeação em 2022. | sim | não | ⚠️ critério |
| E4 | **Afastamento e o VR**: a planilha zera licença previdência/INSS, sem vencimento, vacância, incapacidade/invalidez e pagamento suspenso; o programa faz o mesmo (mês inteiro → zera; parte do mês → desconta os dias úteis afastados). **Licença-maternidade e doenças mantêm.** ⚠️ **Conflito com a lei**: a **Lei 2.729/2019** (não revogada pela 3.035/2023) estende o VR a quem está **afastado pelo INSS**, em licença remunerada por doença de familiar, em licença-maternidade e aos reabilitados, desde que já recebessem antes. A Lei 3.035/2023 só tira o VR em férias, recesso escolar e licenças sem remuneração (art. 78, § 2º). Para seguir a lei, basta tirar "previdencia" e "inss" de **Configurações > Regras > Afastamentos que tiram o VR**. | | | ⚠️ confirmar |
| E5 | Situação "EM APOSENTAÇÃO" trava até alguém verificar. | | | ⚠️ |
| E6 | Quem estava na planilha e **sumiu do SIGA** (exonerado, desligado) sai da planilha, com aviso. Contrato que **termina dentro do mês** (EXONERACAO do SIGA) recebe **proporcional** até a data. | | | ✅ / ⚠️ proporcional |
| E7 | Quem foi **zerado à mão** no mês anterior (ex.: "OPTOU POR RECEBIMENTO TRT", "VACÂNCIA DO CARGO") **continua zerado**, com aviso para reativar se for o caso. Se o zero foi do próprio programa (afastamento ou limite salarial) e a regra deixou de valer, volta a receber com aviso. | | | ✅ planilha |

## 6. Novos admitidos e folha complementar

| # | Regra | Fonte |
|---|---|---|
| N1 | 50 a 70 contratos por mês. Para cada um é preciso confirmar matrícula, **data de admissão**, cargo, jornada e secretaria [06:23, 21:55]. O robô de consulta lê isso na Ábaco (passo 2). | ✅ |
| N2 | Admitido no mês **M** recebe **complementar**: dias úteis da admissão até o fim de M, mais o mês M+1. Entra na planilha normal no ciclo seguinte. Exemplo: admitidos em outubro → complementar de outubro e novembro, entram na planilha do fim de novembro, com crédito em dezembro [28:30, 30:41]; POP VR item 12. | ✅ |
| N3 | Há exceções na complementar [28:30]: o programa calcula e a pessoa confere. Plantonista novo: complementar pelos plantões feitos (fica pendente). | ⚠️ |
| N4 | O VA de novo admitido entra a partir do ciclo seguinte; complementar de VA não está definida. | ⚠️ |
| N5 | O SIGA emite o relatório de novos admitidos por período [25:32]: coluna **NOMEACAO**. | ✅ |
| N6 | **TIPO do novo admitido**: comissionado → NORMAL. Se ≥ 90% dos servidores do mesmo cargo na planilha têm o mesmo tipo (ex.: VIGIA → PLANTÃO), usa esse tipo e avisa. Senão, sugere pela jornada do cargo (40h NORMAL, 30h DIFERENCIADO, 20h DIFERENCIADO 2) e **trava** até a confirmação no arquivo de ajustes. | ⚠️ |

## 7. Educação

| # | Regra | Fonte |
|---|---|---|
| D1 | O **VR** da Educação é feito pela **SEMED** (sábados letivos, recessos) [24:22]; POP VR itens 4 e 8. Nenhum servidor da Educação está na aba CARGA. | ✅ |
| D2 | A Divisão manda à SEMED uma planilha com **matrícula, nome, data de admissão e cargo** dos novos admitidos; a SEMED devolve pronta para importar [24:22]. | ✅ |

## 8. Vale-transporte (VT): próxima fase

| # | Regra | Fonte |
|---|---|---|
| T1 | O problema do VT são as **várias tarifas** [14:32]. | ✅ |
| T2 | **Cada empresa de transporte** usa um formato próprio de importação [15:04]. | ✅ |
| T3 | Ordem combinada: **primeiro VR, depois VT** [32:39]; o VT começa depois que a folha fecha [16:14]. | ✅ |

> O VT **ainda não está no programa**: faltam as tabelas de tarifa e o
> modelo de arquivo de cada empresa.

---

## 9. Arquivos do mês

**Entrada** (pasta `entrada/`; a tela Benefícios guia cada um):

| Arquivo | De onde | Colunas que o programa usa |
|---|---|---|
| `carga_anterior` | planilha do mês passado (no 1º mês, a `11- NOVEMBRO - 2026.xlsx`; depois a que o programa gerou) | aba CARGA: MAT, NOME, TIPO, QTDE, OBSERVAÇÃO |
| `servidores` | SIGA, relatório "Situação de afastamento" | MATRICULA, NOME, CPF, SITUACAO, REGIME, NOME_SECRETARIA, NOMEACAO, EXONERACAO, CARGO_EFETIVO, CARGO_COMISSIONADO, SALARIO, INICIO/TERMINO/TIPO_AFASTAMENTO |
| `ferias` | Ábaco, período de férias (3 meses) | Matricula, Período de Gozo (ou Início/Fim) |
| `faltas` | Ábaco, faltas | Matricula, Qtde, Competência |
| `plantoes` | e-mail da secretaria | matrícula, quantidade de plantões |
| `conferencia_abaco` | robô (passo 2) | admissão, cargo, jornada, secretaria lidos na Ábaco |
| `ajustes` | a própria Divisão | Matrícula, Tipo, Quantidade, Desconto, Reembolso, Observação |

**Saída** (pasta `saida/AAAA-MM/` da competência):

| Arquivo | Para quê |
|---|---|
| `11- NOVEMBRO - 2026.xlsx` (nome do mês da folha) | **a planilha no formato da Divisão**: aba CARGA com as mesmas colunas e fórmulas (QTDE `=$I$3`, VALOR pelo tipo, TOTAL), mais FERIAS, FALTAS e DADOS. Pode ser editada no Excel e é a **base do mês seguinte**. |
| `VR_VA_AAAA-MM_conferencia.xlsx` | conferir: Resumo, Pendências, VR, VA, Complementar, Excluídos, Aguardando, Educação, Fora da planilha |
| `VR_AAAA-MM_importacao_abaco.xlsx` | importar o VR (matrícula, valor) |
| `VA_AAAA-MM_credito_…` / `VA_AAAA-MM_desconto_…` | importar o crédito e o desconto do VA |
| `Complementar_AAAA-MM.xlsx` | complementar dos admitidos em M-2 |
| `Educacao_AAAA-MM_novos_admitidos.xlsx` | enviar à SEMED |
| `ajustes_para_preencher_AAAA-MM.xlsx` | as pendências em forma de ajuste: preencher, salvar como `entrada/ajustes.xlsx` e calcular de novo |

## 10. Conferência com os dados reais (04/11/2026)

O programa rodou com a `11- NOVEMBRO - 2026.xlsx`, o relatório do SIGA e
as abas FERIAS e FALTAS da própria planilha, para a competência 12/2026:

- **2.310 linhas com exatamente o mesmo TOTAL** da planilha.
- As outras diferenças são de dados ainda não atualizados na planilha
  (19 dias de novembro digitados na QTDE; férias e faltas do ciclo
  anterior) ou das regras E4/E6/R8 aplicadas pelo SIGA (afastados pelo
  INSS, contratos que terminam em dezembro, salário acima do limite).
- As fórmulas da planilha gerada, recalculadas no LibreOffice, deram o
  mesmo total do programa nas 2.794 linhas (R$ 3.073.378,00).
- Encontrado e corrigido: o SIGA escreve "AGENTES POLITICOS" (plural), e a
  regra antiga não pegava; 22 agentes políticos receberiam VA.

## 11. Pendências para fechar com a área

- [ ] **Afastado pelo INSS recebe VR?** A Lei 2.729/2019 diz que sim; a planilha zera (E4). É o ponto mais importante.
- [ ] **Decreto vigente** com o valor diário do VR por tipo e o **limite salarial** (R8: R$ 5.662,23 bate com a correção anual, mas o decreto não está na internet). Base: vencimento básico (a lei diz "vencimento").
- [ ] VA de **novo admitido** (N4).
- [ ] **Contrato que termina no meio do mês** recebe proporcional (E6)?
- [ ] **Afastamento em parte do mês** (E4): desconta os dias (como a antiga aba DEDUCAO) ou não?
- [ ] Faltas em **mais de uma linha** para a mesma matrícula: somar (programa) ou só a primeira (planilha)?
- [ ] **Duplo vínculo** no VR (R10): como fica o valor de cada matrícula?
- [ ] Layout exato dos **arquivos de importação** da Ábaco (VR e VA crédito/desconto).
- [ ] Feriados municipais e pontos facultativos do ano (decretos).
- [ ] Tarifas e modelos de arquivo por empresa para o VT.

## 12. O que a legislação confirma (pesquisa de 04/11/2026)

Os decretos 13.655 e 13.719 citados no POP **não estão publicados na
internet** (nem no site da Prefeitura nem no da Câmara). As leis abaixo
estão no site da Câmara Municipal.

| Norma | O que diz | Efeito no programa |
|---|---|---|
| [Lei 3.035, de 13/09/2023](https://cmnovalima.mg.gov.br/processo-legislativo/arquivos/765d737bc94a798dfcfcb5b5d83d4336.pdf) | Regulamenta o acordo judicial com o SINDSERP; reescreve o **art. 78 (VR)** e cria o **art. 79-A (VA)** do Estatuto (Lei 2.590/2017); revoga as Leis 2.727 e 2.728/2019; efeitos desde 01/07/2023. **VR**: jornada < 8h com vencimento até o limite do decreto; 8h/40h e 12x36 independentemente do vencimento; 12x36 só nos dias trabalhados; não é devido em férias, recesso e licenças sem remuneração; cedido com ônus do Município recebe; valor por decreto. **VA**: todos os servidores, qualquer jornada; mantido em ausência justificada, férias, recesso, INSS e licença por doença de familiar; perdido na ausência injustificada e na licença sem remuneração; cedido com ônus do Município recebe; valor por decreto. Anexo (acordo): VR R$ 22,50/dia (8h), R$ 33,75 (12x36), R$ 16,87 (< 8h, vencimento até R$ 2.058,70); VA R$ 400,00 para todos. | R6, R7, R8, A1, A2, A5, A6, E2 confirmados; VA no INSS e VA dos cedidos **corrigidos** |
| [Lei 2.729, de 25/11/2019](https://cmnovalima.mg.gov.br/processo-legislativo/arquivos/483fe0cd256e8bd5130d066b6206c102.pdf) | Estende o VR a afastados pelo INSS, licença remunerada por doença de familiar, licença-maternidade e reabilitados, para quem já recebia; sem o desconto de R$ 1,00 na suspensão. Não foi revogada pela Lei 3.035. | Conflito com a planilha (E4) |
| [Lei 2.728/2019](https://cmnovalima.mg.gov.br/processo-legislativo/arquivos/fa9dde0629ef1c5f75eef999370b2b94.pdf) e [Lei 2.727/2019](https://cmnovalima.mg.gov.br/processo-legislativo/arquivos/c1298e7915ba8c0c79bf8118dcded926.pdf) (revogadas) | Origem do VA por faixa salarial, do desconto de R$ 1,00, da exclusão do aposentado e do limite de vencimento do VR para jornada < 8h com correção anual pelo índice geral. | Histórico do R8 e do A3 |
| [Notícia da Prefeitura, 21/08/2023](https://novalima.mg.gov.br/inicio/noticias/prefeitura_e_sindserp_firmam_acordo_historico) e [Jornal Sempre Nova Lima, 22/08/2023](https://semprenovalima.com/2023/08/prefeitura-e-sindserp-firmam-acordo-que-busca-aumento-dos-vales-refeicao-e-alimentacao-para-servidores/) | Acordo: VA R$ 400,00 para todos; VR R$ 22,50/dia (40h), proporcional para outras jornadas. | A1, R7 |
| [Câmara, 10/07/2026](https://cmnovalima.mg.gov.br/noticias/debate-em-plenario-marca-aprovacao-do-reajuste-dos-servidores-municipais-em-nova-lima/) | PL 2.732/2026: reajuste de 4,2% aprovado em 09/07/2026. | Explica o limite de R$ 5.662,23 (R8) |
| [Ata da Câmara, 17/10/2023](https://cmnovalima.mg.gov.br/processo-legislativo/arquivos/b54abf3a82ac4492d9430487b6a94291.pdf) | Requerimento para reajustar o VR e estendê-lo a jornadas < 8h; o decreto 13.719 (POP) é de dois dias depois. | Contexto do R8 |

Não há base legal pública para: **PSS de 2022** sem VA (só a reunião),
**agente político** sem VA (só o POP), **contrato que termina no mês**
pago proporcional, **layout de importação** da Ábaco. Esses pontos
continuam com a Divisão.

