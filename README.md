# Como funciona — Automação Ábaco + Excel

Programa com tela (Qt6) que faz duas coisas:

1. **Benefícios VR/VA**: o fluxo da Divisão de Benefícios (reunião de
   03/11/2026, POP de VA/VR e a planilha de carga mensal). Parte da
   planilha do mês anterior, como a Divisão faz, atualiza com o SIGA e os
   relatórios de férias, faltas e plantões, calcula vale-refeição,
   vale-alimentação e complementar, e gera a planilha no mesmo formato
   (`11- NOVEMBRO - 2026.xlsx`) mais os arquivos de importação.
2. **Diagnóstico**: entender como uma pessoa trabalha na Ábaco antes de
   automatizar uma tarefa nova.

As regras de negócio, os requisitos e o que ainda falta validar com a
área estão em **[REGRAS_DE_NEGOCIO.md](REGRAS_DE_NEGOCIO.md)**.

> O robô prepara, a pessoa confere. Nada vai para a Ábaco sem a
> conferência de quem é responsável pelo benefício.

---

## Como abrir

**Dois cliques no `INICIAR.bat`.** Ele faz tudo sozinho, em qualquer
computador com Windows 10/11 de 64 bits:

1. confere o ambiente isolado do programa, que fica em
   `%LOCALAPPDATA%\AutomacaoAbaco\.venv` (fora da pasta do programa). Por
   isso dá para copiar a pasta para outro computador ou pen drive: cada
   computador monta o seu ambiente. Se o ambiente estiver faltando ou
   quebrado, é recriado;
2. para criar o ambiente, procura um **Python 3.10 a 3.14 de 64 bits**
   (lançador `py`, registro do Windows, pastas padrão e PATH). Se não
   houver, **instala o Python 3.12 só para este usuário** (não precisa de
   administrador): primeiro pelo `winget` e, se não der, pelo instalador
   oficial do python.org, que só é executado se a assinatura digital for
   da Python Software Foundation. O download usa o proxy do Windows;
3. instala as bibliotecas (Qt6, Playwright, pandas...), confere se elas
   carregam e baixa o navegador controlado (Chromium). Se o Chromium não
   puder ser baixado, o programa usa o **Microsoft Edge** (ou Chrome) do
   próprio Windows. Na primeira vez leva alguns minutos; depois é pulado,
   a não ser que o `requirements.txt` mude;
4. abre a tela do programa. Se ela fechar com erro logo ao abrir, a
   janela mostra o motivo.

Use o mesmo `INICIAR.bat` todo dia. Opções para quem dá suporte:

```
INICIAR.bat -SoInstalar     instala e não abre
INICIAR.bat -Reinstalar     apaga o ambiente e instala tudo de novo
```

---

## As telas

A barra da esquerda leva às telas. Embaixo de todas fica o **Registro**,
que mostra tudo o que o programa está fazendo. Tarefas que usam o
navegador rodam em segundo plano, e a barra de status mostra a tarefa em
andamento com um botão **Cancelar tarefa**.

Quando o robô precisa da pessoa ("faça login", "baixou o relatório?"),
aparece um **painel destacado com botões** acima do registro. A tela
nunca trava.

### Início
Situação do programa: bibliotecas, navegador, endereço da Ábaco, valores
de VR/VA e robô de consulta. Cada item mostra ✔ ou ⚠ com o que fazer.
Também mostra a rotina do mês.

### Benefícios VR/VA
Escolha a **competência**, que é o mês de **uso** do benefício. O padrão é
o mês seguinte, porque o pagamento é antecipado: a competência 12/2026 é
calculada na folha de novembro, como na planilha `11- NOVEMBRO - 2026`.
Depois siga os três passos.

**1. Coletar relatórios.** A lista mostra o que já está na pasta e
quando chegou:

| Arquivo | De onde | O que é |
|---|---|---|
| `carga_anterior` | Divisão | **a planilha do mês anterior** (aba CARGA com o TIPO de cada servidor). No 1º mês, adicione a `11- NOVEMBRO - 2026.xlsx`; depois o programa usa sozinho a que ele gerou |
| `servidores` | SIGA | relatório "Situação de afastamento" de todos os servidores |
| `ferias` | Ábaco | período de férias, janela de 3 meses (sugerida na tela) |
| `faltas` | Ábaco | faltas do mês anterior à folha (sugerido na tela) |
| `plantoes` | e-mail da secretaria | quantidade de plantões por matrícula |
| `ajustes` | Divisão | opcional: correções de tipo, quantidade, desconto e reembolso |

"Coletar relatórios" pede a planilha do mês anterior (se faltar), abre o
navegador e guia cada relatório: a pessoa gera **em Excel** e baixa, e o
programa salva com o nome certo. Se vier em PDF, ele avisa. No fim, pede
a planilha de plantões. Também dá para **adicionar um arquivo
manualmente**. Os relatórios não precisam ter exatamente os nomes de
coluna esperados: o programa reconhece variações e encontra o cabeçalho
mesmo com título acima dele.

**2. Conferir novos admitidos na Ábaco (opcional).** O robô **digita cada
matrícula nova** (admitidos desde M-2 que não estão na planilha) na tela
de consulta do servidor e lê a data de admissão, o cargo, a jornada e a
secretaria. No cálculo, **os dados da Ábaco prevalecem** e cada diferença
vira pendência. Proteção: se o nome lido na tela não bate com o do SIGA,
a linha é descartada. Antes de usar, configure os seletores
(**Configurações > Robô de consulta**).

**3. Calcular VR/VA.** Mostra os **dias úteis** sem os feriados e avisa
sobre pontos facultativos comuns no mês (Carnaval, Corpus Christi...). O
número pode ser ajustado. O cálculo **atualiza a planilha do mês
anterior**, como a Divisão faz:

- **VR**: TOTAL = (dias úteis − férias − faltas) × valor do TIPO
  (NORMAL R$ 56, PLANTÃO R$ 84, DIFERENCIADO R$ 42, DIFERENCIADO 2
  R$ 28), − desconto + reembolso. **Plantonista**: plantões × valor, sem
  dedução; sem a planilha da secretaria, mantém a quantidade do mês
  anterior e avisa.
- **Quem sai ou zera**: quem sumiu do SIGA, exonerados, afastados pelo
  INSS, sem vencimento ou vacância, e diferenciado com salário acima do
  limite (R$ 5.662,23). Quem foi zerado à mão no mês anterior continua
  zerado (com aviso). Contrato que termina no mês recebe proporcional.
- **Novos admitidos** (mês M-2): entram na planilha com **complementar**
  (retroativo desde a admissão + o mês seguinte). O TIPO vem do cargo:
  comissionado → NORMAL; cargo em que quase todos têm o mesmo tipo (ex.:
  vigia → PLANTÃO) → esse tipo; senão, fica **pendente** até alguém
  confirmar no arquivo de ajustes. Admitidos em M-1 aguardam o próximo
  ciclo.
- **Educação**: o VR é feito pela SEMED; os novos admitidos vão para uma
  planilha à parte. No **VA**, a Educação entra.
- **VA**: crédito de R$ 400,00 e desconto de R$ 1,00; **uma falta
  injustificada faz perder o mês** (sem o desconto). Não recebem VA:
  agentes políticos, PSS de 2022, estagiários, aposentados, pensionistas
  e quem está o mês todo em licença sem remuneração. Quem está afastado
  pelo INSS e o cedido com ônus do Município **recebem** o VA (Lei
  3.035/2023). Agente político **recebe VR**.

Os detalhes e a fonte de cada regra estão em
[REGRAS_DE_NEGOCIO.md](REGRAS_DE_NEGOCIO.md).

### Resultados
Abre sozinha depois do cálculo. Traz o total do VR, os números-resumo, os
avisos e uma aba por lista: **Pendências, VR, VA, Complementar,
Excluídos, Aguardando próximo ciclo, Educação e Fora da planilha** (ativos
que não estão na planilha, para conferir). As linhas pendentes ficam **em
laranja**, e dá para filtrar e ordenar.

Arquivos gerados em `saida/AAAA-MM/` (AAAA-MM = competência):

| Arquivo | Para quê |
|---|---|
| `11- NOVEMBRO - 2026.xlsx` (mês da folha) | **a planilha da Divisão**: aba CARGA com as mesmas colunas e fórmulas (dá para mexer no Excel), mais FERIAS, FALTAS e DADOS. É a **base do mês seguinte** |
| `VR_VA_AAAA-MM_conferencia.xlsx` | **conferir**: as mesmas abas, com a memória de cálculo de cada linha |
| `VR_AAAA-MM_importacao_abaco.xlsx` | importar o VR na Ábaco (matrícula e valor; **só linhas sem pendência e com valor**) |
| `VA_AAAA-MM_credito_importacao_abaco.xlsx` | importar o crédito do VA |
| `VA_AAAA-MM_desconto_importacao_abaco.xlsx` | importar o desconto de R$ 1,00 do VA |
| `Complementar_AAAA-MM.xlsx` | folha complementar dos novos admitidos |
| `Educacao_AAAA-MM_novos_admitidos.xlsx` | enviar à Secretaria de Educação |
| `ajustes_para_preencher_AAAA-MM.xlsx` | as pendências prontas para corrigir |

**Como resolver uma pendência**: abra o `ajustes_para_preencher`,
preencha só o que muda (Tipo, Quantidade, Desconto, Reembolso),
salve como `ajustes.xlsx` na pasta de entrada (ou use "Adicionar arquivo
manualmente" > Ajustes) e calcule de novo. Linhas que "travam a
importação" **não entram** nos arquivos de importação até serem
resolvidas.

### Diagnóstico
- **Observar o trabalho**: abre a Ábaco, e a pessoa trabalha normalmente
  enquanto o programa registra o **processo** (telas, cópias e colagens
  e de qual campo). Há um painel **"🔴 Observando"** na tela o tempo todo.
  Ao fechar o navegador, sai o `relatorio_observacao.md`.
- **Gravar fluxo**: a pessoa executa uma tarefa **uma vez** e o gravador
  do Playwright escreve o passo a passo em `fluxo_gravado.py`. É assim que
  se descobrem os seletores do robô de consulta.
- **Extrair tabela para Excel**: lê a tabela que está na tela da Ábaco,
  em várias páginas seguidas, e grava num `.xlsx`.

### Configurações
Formulários, salvos no `config.json`:

| Aba | O que ajusta |
|---|---|
| Geral | endereço da Ábaco, pastas, navegador visível |
| Valores | valor por dia de cada tipo (NORMAL, PLANTÃO, DIFERENCIADO, DIFERENCIADO 2), limite salarial do diferenciado, plantões de um plantonista novo, crédito e desconto do VA, regra do VA com falta, VA para a Educação |
| Calendário | feriados municipais e pontos facultativos (os nacionais são automáticos) |
| Regras | lotação da Educação, afastamentos que tiram o direito, situações, quem não recebe VR e quem não recebe VA, prazos |
| Importação | colunas dos arquivos importados na Ábaco (VR, VA crédito, VA desconto) |
| Robô de consulta | seletores da tela de consulta do servidor (ex.: `#txtMatricula`) |
| Colunas | nomes de coluna diferentes usados nos relatórios |

Valores e datas são conferidos ao salvar (ex.: uma data 31/02 não passa).

### Instalação
Mostra a versão do Python, cada biblioteca e o navegador, e reinstala as
bibliotecas e o navegador com um clique.

---

## Rotina do mês (benefícios)

1. Até o **dia 5**: chega a documentação; até o dia 10, os contratos dos
   novos admitidos; a secretaria envia os plantões.
2. Até o **dia 15**: **coletar** os relatórios da competência (SIGA,
   férias e faltas).
3. **Conferir** os novos admitidos na Ábaco (quando o robô estiver
   configurado).
4. **Calcular**, revisar em **Resultados** e resolver as pendências pelo
   arquivo de ajustes.
5. Importar VR e VA (crédito e desconto) na Ábaco **até o dia 20**
   (fechamento da folha).
6. Enviar a planilha da Educação à SEMED.

---

## O que é e o que NÃO é registrado na observação  *(importante)*

A observação foi feita para **diagnóstico de processo**, não para vigiar
uma pessoa.

**Registra:** troca de telas (título e endereço), que houve uma cópia
(de qual campo e quantos caracteres), que houve uma colagem (em qual
campo) e as anotações que a pessoa escolhe escrever.

**NÃO registra:** o **conteúdo** copiado (nada de CPF, valores ou nomes),
nada fora do navegador que o programa abriu, nem a **senha** da Ábaco.
Em todas as funções, a senha é digitada pela pessoa na tela oficial do
sistema e nunca passa pelo programa nem é guardada.

---

## Testes

As regras de VR/VA têm testes com dados fictícios (incluem os exemplos
citados na reunião: 19 dias úteis em novembro/2026 e 22 em dezembro):

```
"%LOCALAPPDATA%\AutomacaoAbaco\.venv\Scripts\python" -m unittest test_regras_beneficios -v
```

---

## Arquivos da pasta

| Arquivo | Para que serve |
|---|---|
| `INICIAR.bat` | **abre o programa** (instala o que faltar, inclusive o Python) |
| `instalar.ps1` | o instalador chamado pelo `INICIAR.bat` |
| `abaco_automacao.py` | a tela (Qt6) |
| `nucleo_abaco.py` | o que o programa faz: navegador, observação, coleta, robô |
| `regras_beneficios.py` | cálculo de VR/VA (não acessa a Ábaco; só lê planilhas) |
| `test_regras_beneficios.py` | testes das regras |
| `REGRAS_DE_NEGOCIO.md` | regras (reunião, POP e planilha), requisitos e pendências |
| `requirements.txt` | bibliotecas instaladas no ambiente do programa |
| `config.json` | criado na 1ª vez; configurações e regras |
| `entrada/` | relatórios do mês, a planilha do mês anterior e os ajustes |
| `saida/AAAA-MM/` | a planilha da Divisão (base do mês seguinte), conferência e importação |
| `modelos/` | planilhas de entrada vazias |
| `relatorio_observacao.md` / `.json` | saída da observação |
| `fluxo_gravado.py` | saída da gravação de fluxo |

---

## Limitações conhecidas

- **VT ainda não está no programa**: falta receber as tarifas e o modelo
  de arquivo de cada empresa de transporte.
- Os valores vieram da planilha e do POP (VR R$ 56/84/42/28 por dia, VA
  R$ 400,00 e R$ 1,00). O **limite salarial vigente** do diferenciado e o
  **layout exato dos arquivos de importação** da Ábaco ainda precisam ser
  confirmados pela área (veja as pendências em `REGRAS_DE_NEGOCIO.md`).
- O robô de consulta depende dos seletores da tela da Ábaco. Se a tela
  estiver dentro de um *iframe* ou mudar de layout, é preciso gravar de
  novo.
- O instalador é para Windows 10/11 (64 bits).
- Se o computador bloqueia scripts do PowerShell ou programas em
  `%LOCALAPPDATA%` (política do TI), o `INICIAR.bat` mostra o erro e é
  preciso pedir a liberação ao suporte.
