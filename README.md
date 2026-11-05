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

### Programa sem Python (AutomacaoAbaco.exe)

Para usar em computadores **sem Python instalado**, gere o programa
empacotado. Faça isso num computador que já abriu pelo `INICIAR.bat`:

1. Dois cliques no **`empacotar.bat`**. Ele confere o ambiente, instala o
   PyInstaller, gera o programa e testa. Leva de 3 a 6 minutos.
2. O resultado é **um arquivo só**: **`dist\AutomacaoAbaco.exe`**. O Python,
   as bibliotecas, o motor do navegador e o **`config.json`** desta pasta
   vão **dentro** dele (não há mais pasta `_internal`).
3. Para outro computador, copie **só o `AutomacaoAbaco.exe`** para a **Área
   de Trabalho ou Documentos** e abra. Não precisa instalar nada.

Como funciona por dentro e o que muda:

- Ao abrir, o `.exe` se descompacta numa pasta temporária do Windows e a
  apaga ao fechar. Por isso **demora alguns segundos a mais para abrir** do
  que a versão em pasta.
- Na 1ª vez, ele cria **ao lado dele** o `config.json` (cópia do embutido) e,
  quando usadas, as pastas `entrada\` e `saida\`. Depois disso vale o
  `config.json` de fora: o que você mudar em Configurações fica salvo nele.
  Para voltar ao embutido, apague o `config.json` de fora.
- Mudou o `config.json` desta pasta (endereço, valores, regras) e quer
  que vá para dentro do `.exe`? Rode o `empacotar.bat` de novo.
- O fluxo gravado (`fluxo_gravado.py`, com a senha) **não** vai dentro do
  `.exe`: grave no próprio computador (Diagnóstico > Gravar fluxo) ou copie
  o arquivo para o lado do `.exe`.
- O robô usa o **Microsoft Edge** do Windows. O Chromium é opcional
  (tela Instalação > "Baixar o navegador controlado", uns 300 MB, em
  `%LOCALAPPDATA%\ms-playwright`). O download usa o **proxy** e os
  **certificados do Windows**; se a rede bloquear, o registro explica o
  motivo (certificado, proxy, sem conexão, pasta bloqueada) e o robô segue
  com o Edge. Endereços que o TI precisa liberar: `cdn.playwright.dev` e
  `playwright.download.prss.microsoft.com`.
- `VERIFICAR.bat`, se colocado ao lado do `.exe`, confere se o programa
  funciona naquele computador sem abrir a tela (grava `verificacao.txt`).
- Não coloque o `.exe` em `C:\Arquivos de Programas`: lá ele não consegue
  gravar o `config.json` e as planilhas.
- Se o antivírus barrar o `.exe`, peça ao TI para liberá-lo (é comum com
  programas gerados pelo PyInstaller, e um pouco mais com os de arquivo
  único).
- A pasta antiga `dist\AutomacaoAbaco\` (com `_internal`) não é mais usada;
  copie dela o que for seu (fluxos, saída) antes de apagá-la.

---

## As telas

O visual segue o arquivo do Figma **"Automação - Fluxos RH (telas)"**
(https://www.figma.com/design/AwhN1qt75rIHVIc7azbUdC): fonte IBM Plex Sans
(e IBM Plex Mono no Registro), embutida em `fontes/`; cada setor tem a sua
cor (Benefícios em azul, Consulta de Admissões em areia). As cores e a folha
de estilo ficam em `tema.py`.

O programa abre na **tela inicial** ("Automação - Fluxos RH"), com os dois
setores. Cada setor mostra **só as funções dele**:

| Na tela inicial | O que aparece |
|---|---|
| **Benefícios VR/VA** (azul) | Benefícios VR/VA (os 3 passos do mês), Resultados e Visão geral |
| **Consulta de Admissões** (bege) | a tela Consulta de admissões |
| **Engrenagem** (canto de baixo) | **Modo desenvolvedor**: todas as funções do programa, agrupadas (Benefícios VR/VA, Consulta de Admissões e Ferramentas: Configurações, Instalação, Diagnóstico), inclusive o robô antigo de conferência na Ábaco e a configuração do robô |

O botão **← Início**, no alto, volta para a tela inicial (uma tarefa em
andamento continua rodando).

Dentro do setor, a barra da esquerda leva às telas dele. Embaixo de todas fica o **Registro**,
que mostra tudo o que o programa está fazendo. Tarefas que usam o
navegador rodam em segundo plano, e a barra de status mostra a tarefa em
andamento com um botão **Cancelar tarefa**.

Quando o robô precisa da pessoa ("faça login", "baixou o relatório?"),
aparece um **painel destacado com botões** acima do registro. A tela
nunca trava.

### Visão geral (setor Benefícios)
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

### Consulta de admissões

Faz em lote o que o `fluxo_consulta_admissao.py` fazia para uma matrícula:
entra no e-Turmalina (sozinho, se a senha está no fluxo gravado), abre
**GESTÃO DE PESSOAS > Dados Funcionais do Servidor**, digita cada matrícula,
clica em **Consultar**, lê a linha da grade (Matrícula, Cnt, Servidor, Cargo)
e abre o ícone **Eventos funcionais do servidor** para ler a ficha (Dados da
nomeação, Dados do cargo: admissão, jornada, lotação...). **Só consulta.**

- **Matrículas**: uma por linha; `22769;2` consulta só o contrato 2. Sem o
  contrato, lê todos os contratos da matrícula. Botões:
  - **Puxar matrículas do SIGA**: o robô entra no SIGA
    (`https://mgnl.abaco.com.br/siga/`), abre **Recursos Humanos > Relação
    de Servidores Admitidos no Período** (visão salva 245), preenche DATA
    ADMISSÃO INICIAL = dia 1º do **mês atual** e FINAL = **hoje** (o dia em
    que o botão é usado: em 25/12, de 01/12 a 25/12), clica em Visualizar
    Resultado e exporta em **XLSX(SEM FORMATO)** (se a exportação não baixar,
    lê a tabela da tela, página por página). O arquivo fica em
    `entrada\siga_admitidos_AAAA-MM-DD_a_AAAA-MM-DD.xlsx`. Entram na lista os
    admitidos que **não estão na planilha base**; os da **Educação** ficam de
    fora (vão para a SEMED). O registro mostra quantos ficaram de fora e por quê.
    **Faz também o que o programa da pasta "003 - GILSON" faz**: grava a
    `admissoes.xlsx` (na pasta do programa) no mesmo formato. A aba
    **Planilha** recebe MAT, NOME, SITUAÇÃO, TIPO, TOTAL DEDUÇÕES, QTDE, VALOR,
    DESCONTO, TOTAL, REEMBOLSO e OBSERVAÇÃO, com cabeçalho Segoe UI azul e
    filtro. Cada admitido entra nas próximas linhas vazias com TIPO NORMAL,
    TOTAL DEDUÇÕES 0, QTDE 19, VALOR 56, DESCONTO 0, TOTAL `=F*G-(E*G)` e
    REEMBOLSO 0. A aba Admissoes_Extraidas é apagada, a planilha abre no Excel
    na aba Planilha com as linhas novas selecionadas, e todos os admitidos do
    SIGA entram nela. Diferença: quem já está na planilha não é repetido (o
    programa do Gilson repete a cada clique). Se a admissoes.xlsx estiver
    aberta no Excel, o botão avisa antes de abrir o navegador. Nome do
    arquivo, QTDE, VALOR e abrir no Excel ficam em Configurações > Geral;
  - **Gravar login do SIGA**: uma vez só. Faça o login e feche o navegador; o
    login (com a senha) fica em `fluxo_siga.py` e o robô entra sozinho. Outra
    opção, como no programa do Gilson: escrever **Usuário e Senha do SIGA** em
    Configurações > Geral (ficam no config.json). Sem nenhum dos dois, o robô
    espera você fazer o login no navegador que abrir;
  - **Do relatório do SIGA na entrada**: usa o relatório "situação de
    afastamento" que já está em `entrada\servidores` (sem abrir o SIGA);
  - **Carregar de planilha...**: a coluna MAT/Matrícula de qualquer aba.
- **Resultado**: `saida\consultas\consulta_admissoes_AAAA-MM-DD_HHMM.xlsx`,
  com 3 abas:
  - **CARGA** (abre nela): **o mesmo formato da aba CARGA da planilha da
    Divisão** — cabeçalho INICIO/FIM/FERIADOS, TOTAL (dias úteis) em I3,
    valores por tipo em K2:L5, tabela `Tabela3` a partir da linha 7 com MAT,
    NOME, SITUAÇÃO, TIPO, (DEDUÇÃO ATESTATO, DE FALTAS e DE FÉRIAS ocultas),
    TOTAL DEDUÇÕES, QTDE (laranja), VALOR, DESCONTO, TOTAL, REEMBOLSO,
    OBSERVAÇÃO e Coluna1, Calibri 8, moeda contábil e as **mesmas fórmulas**.
    O TIPO vem da jornada semanal (200/40 → 40 h → NORMAL) e a OBSERVAÇÃO diz
    a admissão e a jornada; confira o limite salarial do DIFERENCIADO e a
    complementar. Dá para copiar as linhas e colar na CARGA do mês;
  - **Consultas**: uma linha por matrícula e contrato, com todos os campos
    lidos (`Ficha: ...` e `Grade: ...`); linhas com erro em laranja;
  - **Fichas**: o texto de cada ficha, para conferir.
- **Usar no cálculo do VR** (marcado): grava também
  `entrada\conferencia_abaco.xlsx`. No cálculo, admissão, cargo, jornada e
  secretaria da Ábaco prevalecem sobre o SIGA e cada diferença vira pendência.
  Com dois contratos, vale o admitido por último.
- Se uma matrícula der erro (não existe, a ficha não abriu), o robô anota e
  segue para a próxima. Se cancelar no meio, o Excel sai com o que já foi lido.

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
  se descobrem os seletores do robô de consulta. Renomeie o arquivo depois
  de cada gravação (ex.: `fluxo_frequencias.py`), senão a próxima o substitui.
- **Reproduzir fluxo...**: escolhe um fluxo gravado e o repete num
  navegador visível, para testar a gravação. Roda **rápido**; marque
  **"Reproduzir devagar (para acompanhar)"** para ver passo a passo (pausa de
  0,6 s entre as ações). Se a senha ficou gravada no fluxo, o robô **faz o
  login sozinho**; se não (ou se a senha mudou), a pessoa entra à mão e o
  robô **começa sozinho** assim que ela entra. Cliques soltos dados na tela
  de login são ignorados.
  Se um item não aparece (submenu que só abre passando o mouse, como
  "Consultas Adicionais por Servidor ▸", ou menu que mudou de lugar), o robô
  procura sozinho e **aprende o caminho**: grava `<fluxo>_corrigido.py` com
  os passos certos (`passar_mouse`, `clicar`). Nas próximas vezes o programa
  usa o corrigido sozinho e vai direto. Se você gravar o fluxo de novo, o
  corrigido antigo deixa de ser usado. Se mesmo assim parar, salva um print
  em `reproducao\`. No fim, o navegador fica aberto até a pessoa fechá-lo.
  Downloads vão para a pasta `reproducao\`. Age no **sistema real**: use só
  com fluxos de consulta e relatório (o programa avisa se o fluxo tiver
  botões como Importar, Confirmar ou Salvar).
- **Ver o código do fluxo**: abre o arquivo no Bloco de Notas, só para ler
  (não reproduz; para isso use Reproduzir fluxo).
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
| `empacotar.bat` | gera o programa que roda **sem Python**, em um arquivo só (`dist\AutomacaoAbaco.exe`) |
| `tema.py` | cores, fontes e folha de estilo do Qt (o visual do Figma) |
| `fontes\` | IBM Plex Sans e Mono (licença SIL OFL), embutidas no programa |
| `AutomacaoAbaco.spec` | receita do PyInstaller usada pelo `empacotar.bat` |
| `VERIFICAR.bat` | vai junto com o .exe: confere o programa sem abrir a tela |
| `instalar.ps1` | o instalador chamado pelo `INICIAR.bat` |
| `abaco_automacao.py` | a tela (Qt6) |
| `nucleo_abaco.py` | o que o programa faz: navegador, observação, coleta, robô |
| `regras_beneficios.py` | cálculo de VR/VA (não acessa a Ábaco; só lê planilhas) |
| `test_regras_beneficios.py` | testes das regras |
| `REGRAS_DE_NEGOCIO.md` | regras (reunião, POP e planilha), requisitos e pendências |
| `requirements.txt` | bibliotecas instaladas no ambiente do programa |
| `config.json` | criado na 1ª vez; configurações e regras |
| `entrada/` | relatórios do mês, a planilha do mês anterior e os ajustes |
| `dist/` | criada pelo `empacotar.bat`: o programa empacotado e o .zip |
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
