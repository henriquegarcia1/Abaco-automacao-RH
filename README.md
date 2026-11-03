# Como funciona — Automação Ábaco + Excel

Programa único e interativo (`abaco_automacao.py`) para, **nesta ordem**:
primeiro **entender** como uma pessoa trabalha na Ábaco, e só depois
**automatizar** o que realmente valer a pena.

---

## A ideia em uma frase

> Antes de construir qualquer robô, a gente observa e documenta o
> trabalho real. O relatório dessa observação é o que diz **se** e
> **o que** vale automatizar.

Automatizar sem olhar o processo é arriscado: você pode gastar tempo
com uma tarefa que é rara, que muda toda hora, ou que nem era gargalo.
Por isso o menu segue a ordem Instalar → **Observar** → Gravar → Extrair.

---

## Como abrir

No computador de quem vai usar, com Python instalado:

```
python abaco_automacao.py
```

Aparece um menu com 5 opções. Da primeira vez, siga na ordem.

---

## O menu, opção por opção

### [1] Instalar  *(uma vez só)*
Instala as bibliotecas (`playwright`, `pandas`, `openpyxl`) e baixa o
navegador controlado (Chromium). Só precisa rodar uma vez por computador.

### [2] OBSERVAR  *(o primeiro passo de verdade)*
Abre o navegador na Ábaco. A pessoa **faz login e trabalha normalmente**.
Enquanto isso, o programa vai registrando o **processo**:

- quais **telas** ela abre e em que ordem;
- **quando** ela copia dados, **de qual campo** e **o tamanho** do que
  foi copiado (ver "O que é e o que não é registrado", abaixo);
- **anotações** que a própria pessoa escreve num painel visível.

Há um painel fixo escrito **"🔴 Observando (pode anotar)"** no canto da
tela o tempo todo — nada acontece escondido. A pessoa digita ali o que
está fazendo ("agora copio o valor pra planilha") e clica em *Anotar*.

Quando ela **fecha o navegador**, o programa gera o relatório (abaixo).

### [3] Gravar fluxo  *(depois de decidir o que automatizar)*
Para uma tarefa específica que você decidiu automatizar. A pessoa executa
essa tarefa **uma vez**, e o programa escreve sozinho o passo a passo em
código (`fluxo_gravado.py`), usando a ferramenta oficial `playwright
codegen`. Esse código é a base do robô final.

### [4] Extrair p/ Excel  *(a execução)*
O caso mais comum aqui: **puxar dados da tela da Ábaco para uma planilha**.
A pessoa navega até a tela com os dados (listagem, relatório), tecla ENTER,
e o programa lê a tabela e grava num `.xlsx`. Dá para ler várias páginas
em sequência.

### [5] Configurações
Ajusta e salva (em `config.json`):

| Item | O que é |
|---|---|
| URL da Ábaco | endereço de login do sistema |
| Arquivo do fluxo | onde o fluxo gravado é salvo |
| Planilha | nome do `.xlsx` de saída |
| Nome do relatório | nome dos arquivos de observação |
| Mostrar navegador | ver o navegador trabalhando (sim) ou rodar oculto (não) |

---

## O que a observação gera

Ao fechar o navegador, saem **dois arquivos**:

- **`relatorio_observacao.md`** — leitura fácil, com:
  - resumo (quantos eventos, cópias, telas, anotações);
  - telas visitadas e quantas vezes;
  - **de onde mais se copiou** (os candidatos a automação);
  - as anotações da pessoa, em ordem;
  - uma **leitura inicial** automática (ex.: *"o campo X foi copiado 3x —
    repetição é o principal sinal de que automatizar vale a pena"*);
  - a linha do tempo completa dos eventos.
- **`relatorio_observacao.json`** — os mesmos dados, em formato bruto,
  para análise mais detalhada depois.

É esse `.md` que orienta a conversa sobre o que automatizar.

---

## O que é e o que NÃO é registrado  *(importante)*

Este programa foi feito para **diagnóstico de processo**, não para
vigiar uma pessoa. Por isso:

**Registra (o padrão do trabalho):**
- a troca de telas (título e endereço da página);
- que houve uma cópia, **de qual campo** e **quantos caracteres**;
- que houve uma colagem e **em qual campo**;
- as anotações que a pessoa escolhe escrever.

**NÃO registra:**
- o **conteúdo** do que foi copiado (nada de CPF, valores, nomes de
  contribuintes) — só o tamanho;
- nada fora do navegador que o programa abriu (não lê o resto do
  computador, não lê outros programas, não lê a planilha do Excel);
- a **senha** da Ábaco — ela é digitada por você na tela oficial do
  sistema e nunca passa pelo programa nem é guardada.

E tudo é visível: o painel "🔴 Observando" fica na tela durante toda a
sessão. Quem opera é quem abre e fecha o programa.

---

## Fluxo de trabalho recomendado

1. **[1] Instalar** (uma vez).
2. Ajustar a **URL da Ábaco** em **[5] Configurações**.
3. **[2] Observar** durante uma ou duas sessões de trabalho real.
4. Analisar o **`relatorio_observacao.md`** e decidir o que vale automatizar.
5. Para cada tarefa escolhida: **[3] Gravar fluxo** e/ou **[4] Extrair**.
6. Ajustar o robô final com base no fluxo gravado.

---

## Requisitos

- Python 3.9 ou mais novo no computador de quem vai usar.
- Acesso normal da pessoa à Ábaco (o login é feito por ela).
- Conexão com a internet (para a opção [1] instalar baixar o navegador).

---

## Arquivos da pasta

| Arquivo | Para que serve |
|---|---|
| `abaco_automacao.py` | o programa (abra com `python abaco_automacao.py`) |
| `config.json` | criado na 1ª vez; guarda as configurações |
| `relatorio_observacao.md` / `.json` | saída da observação |
| `fluxo_gravado.py` | saída da gravação de fluxo |
| `dados.xlsx` | saída da extração para Excel |
| `COMO-FUNCIONA.md` | este arquivo |

---

## Limitações conhecidas

- O leitor de tabela da opção [4] cobre tabelas HTML comuns. Se a Ábaco
  exibir os dados numa grade própria, basta um print da tela para ajustar
  o leitor ao layout certo.
- A observação enxerga apenas o que acontece **dentro** do navegador da
  Ábaco. A colagem que acontece no Excel (fora do navegador) não é vista —
  mas a **cópia** correspondente, feita na Ábaco, é registrada, e é ela
  que marca o ponto de passagem Ábaco → Excel.
