# Diagnóstico > Reproduzir fluxo. Faça login no navegador quando solicitado.
# O Excel individual fica na pasta reproducao, ao lado da planilha de origem.


def run(playwright):
    from pathlib import Path
    from time import monotonic
    from datetime import datetime
    import base64
    import os
    import subprocess
    from tempfile import TemporaryDirectory
    from urllib.parse import urljoin
    from playwright.sync_api import TimeoutError as TimeoutErrorPlaywright
    from nucleo_abaco import PASTA
    from regras_beneficios import norm
    matricula = "22771"
    contrato = "1"
    pastas = (PASTA, PASTA.parent, PASTA.parent.parent)
    original = next((arq for pasta in pastas for arq in pasta.glob("12*2026.xlsx") if re.fullmatch(r"12\s*-\s*DEZEMBRO\s*-\s*2026\.xlsx", arq.name, re.I)), None)
    base = original or next((pasta / "11- NOVEMBRO - 2026.xlsx" for pasta in pastas if (pasta / "11- NOVEMBRO - 2026.xlsx").is_file()), None)
    if base is None:
        raise RuntimeError("Coloque a planilha de dezembro ou a base de novembro na pasta do programa.")

    def aguardar(procurar, mensagem, segundos=20):
        limite = monotonic() + segundos
        while monotonic() < limite:
            achados = []
            for aba in context.pages:
                for quadro in aba.frames:
                    achado = procurar(quadro)
                    if achado is not None:
                        achados.append(achado)
            if len(achados) == 1:
                return achados[0]
            if len(achados) > 1:
                raise RuntimeError("A tela mostrou mais de uma correspondência. " + mensagem)
            page.wait_for_timeout(250)
        raise RuntimeError(mensagem)

    def formulario(quadro):
        botao = quadro.get_by_role("button", name="Consultar", exact=True)
        if botao.count() != 1 or not botao.is_visible():
            return None
        rotulo = re.compile(r"^\s*Matr[íi]cula\s*:?\s*$", re.I)
        campos = (
            quadro.get_by_label(rotulo),
            quadro.get_by_text(rotulo).locator("xpath=ancestor::td[1]/following-sibling::td[1]").locator("input:visible"),
            quadro.locator("input[id*='matric' i]:visible, input[name*='matric' i]:visible"),
        )
        for campo in campos:
            if campo.count() == 1 and campo.is_visible() and campo.is_editable():
                return campo, botao
        return None

    def resultado(quadro):
        cabecalhos = quadro.get_by_text(re.compile(r"^\s*Cnt\s*\.?\s*$", re.I)).locator("xpath=ancestor::tr[1]")
        grades = []
        for cabecalho in cabecalhos.all():
            colunas = [norm(c) for c in cabecalho.locator(":scope > th, :scope > td").all_inner_texts()]
            if {"matricula", "cnt", "servidor", "cargo"}.issubset(colunas):
                grades.append((cabecalho.locator("xpath=ancestor::table[1]"), colunas))
        if not grades:
            return None

        linhas = quadro.locator("tr:visible").filter(has=quadro.get_by_text(matricula, exact=True))
        achados = []
        for linha in linhas.all():
            valores = [" ".join(v.split()) for v in linha.locator(":scope > td").all_inner_texts()]
            if matricula not in valores:
                continue
            tabela = linha.locator("xpath=ancestor::table[1]")
            compativeis = [colunas for grade, colunas in grades if tabela.and_(grade).count()]
            if not compativeis and len(grades) == 1:
                compativeis = [grades[0][1]]
            for colunas in compativeis:
                dados = dict(zip(colunas, valores))
                if dados.get("matricula") == matricula and dados.get("cnt") == contrato and dados.get("servidor"):
                    achados.append((quadro, linha, dados["servidor"], colunas.index("cargo")))
        if len(achados) > 1:
            raise RuntimeError("A consulta retornou mais de uma linha para a matrícula e o vínculo do vídeo.")
        return achados[0] if achados else None

    def acao_funcional(quadro, linha, indice_cargo):
        rotulo = "Eventos funcionais do servidor"
        atributos = ("title", "alt", "aria-label", "data-tooltip", "onmouseover", "onmouseenter")
        seletor = ", ".join(f"[{atributo}*='{rotulo}' i]:visible" for atributo in atributos)
        icones = linha.locator(seletor)
        acionaveis = icones.locator("xpath=ancestor-or-self::*[self::a or self::button or @role='button'][1]")
        acao = acionaveis if acionaveis.count() else icones
        if acao.count() == 1:
            return acao

        # A grade do vídeo exibe quatro ícones após Cargo. Confirme pelo tooltip.
        celulas = linha.locator(":scope > td")
        candidatos = []
        for indice in range(indice_cargo + 1, celulas.count()):
            celula = celulas.nth(indice)
            controle = celula.locator("a:visible, button:visible, [role=button]:visible, input[type=image]:visible")
            if controle.count() == 0:
                controle = celula.locator("img:visible, svg:visible")
            if controle.count() == 1:
                candidatos.append(controle)
        if len(candidatos) == 4:
            candidatos = [candidatos[1], candidatos[0], *candidatos[2:]]
        dica = quadro.get_by_text(re.compile(r"^\s*Eventos\s+funcionais\s+do\s+servidor\s*$", re.I)).filter(visible=True)
        for candidato in candidatos:
            if dica.count():
                break
            try:
                candidato.hover(timeout=1200)
                dica.first.wait_for(state="visible", timeout=1500)
                return candidato
            except TimeoutErrorPlaywright:
                continue
        return None

    def ficha(quadro):
        corpo = quadro.locator("body")
        if not corpo.count():
            return None
        texto = corpo.inner_text(timeout=1500)
        if "dados da nomeacao" in norm(texto) and "dados do cargo" in norm(texto):
            return texto
        return None

    def gravar_admissao(base, destino, nome, quantidade):
        # Recalcula na base e exporta somente os valores da linha consultada.
        comando = r"""
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
        $excel = $null
        $livro = $null
        $livroSaida = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.EnableEvents = $false
            $excel.AutomationSecurity = 3
            $livro = $excel.Workbooks.Open($env:ABACO_ORIGEM, 0, $true)
            $aba = $livro.Worksheets.Item('CARGA')
            $tabelas = @()
            for ($i = 1; $i -le $aba.ListObjects.Count; $i++) {
                $candidata = $aba.ListObjects.Item($i)
                $nomes = @($candidata.ListColumns | ForEach-Object { $_.Name })
                if ($nomes -contains 'MAT' -and $nomes -contains 'NOME' -and
                    $nomes -contains 'TIPO' -and $nomes -contains 'QTDE') {
                    $tabelas += $candidata
                }
            }
            if ($tabelas.Count -ne 1) {
                throw 'A aba CARGA precisa ter uma tabela identificada por MAT, NOME, TIPO e QTDE.'
            }
            $tabela = $tabelas[0]
            $colMat = $tabela.ListColumns.Item('MAT').Index
            $colNome = $tabela.ListColumns.Item('NOME').Index
            $colTipo = $tabela.ListColumns.Item('TIPO').Index
            $colQtde = $tabela.ListColumns.Item('QTDE').Index
            $indices = @()
            if ($tabela.ListRows.Count -gt 0) {
                $valores = $tabela.ListColumns.Item('MAT').DataBodyRange.Value2
                for ($i = 1; $i -le $tabela.ListRows.Count; $i++) {
                    $valor = if ($valores -is [array]) { $valores[$i, 1] } else { $valores }
                    if ([string]$valor -eq $env:ABACO_MAT) { $indices += $i }
                }
            }
            if ($indices.Count -gt 1) {
                $originais = @($indices | Where-Object {
                    $tabela.ListRows.Item($_).Range.Cells.Item(1, $colNome).Value2
                })
                if ($originais.Count -ne 1) {
                    throw 'As linhas repetidas têm dados ambíguos. Revise a CARGA.'
                }
                $principal = $originais[0]
                $extras = @($indices | Where-Object { $_ -ne $principal } | Sort-Object -Descending)
                foreach ($i in $extras) {
                    $linhaExtra = $tabela.ListRows.Item($i).Range
                    for ($j = 1; $j -le $tabela.ListColumns.Count; $j++) {
                        if ($j -eq $colMat) { continue }
                        $celula = $linhaExtra.Cells.Item(1, $j)
                        if (-not $celula.HasFormula -and [string]$celula.Value2 -ne '') {
                            throw 'Uma linha repetida contém dados manuais. Revise a CARGA.'
                        }
                    }
                }
                foreach ($i in $extras) {
                    $tabela.ListRows.Item($i).Delete()
                }
                $indices = @($principal - @($extras | Where-Object { $_ -lt $principal }).Count)
            }
            if ($indices.Count -eq 1) {
                $linha = $tabela.ListRows.Item($indices[0]).Range
            } else {
                $anterior = $null
                if ($tabela.ListRows.Count -gt 0) {
                    $anterior = $tabela.ListRows.Item($tabela.ListRows.Count).Range
                }
                $linha = $tabela.ListRows.Add().Range
                if ($null -ne $anterior) {
                    for ($j = 1; $j -le $tabela.ListColumns.Count; $j++) {
                        $modelo = $anterior.Cells.Item(1, $j)
                        $nova = $linha.Cells.Item(1, $j)
                        if ($modelo.HasFormula -and -not $nova.HasFormula) {
                            $nova.FormulaR1C1 = $modelo.FormulaR1C1
                        }
                    }
                }
            }
            $linha.Cells.Item(1, $colMat).Value2 = [double]$env:ABACO_MAT
            $linha.Cells.Item(1, $colNome).Value2 = $env:ABACO_NOME
            $linha.Cells.Item(1, $colTipo).Value2 = 'NORMAL'
            if ($env:ABACO_QTDE) {
                $linha.Cells.Item(1, $colQtde).Value2 = [double]$env:ABACO_QTDE
            }
            $aba.Calculate()
            $livroSaida = $excel.Workbooks.Add(-4167)
            $abaSaida = $livroSaida.Worksheets.Item(1)
            $abaSaida.Name = 'CARGA'
            $colunas = @('MAT', 'NOME', 'SITUAÇÃO', 'TIPO', 'TOTAL DEDUÇÕES',
                         'QTDE', 'VALOR', 'DESCONTO', 'TOTAL', 'REEMBOLSO', 'OBSERVAÇÃO')
            for ($j = 0; $j -lt $colunas.Count; $j++) {
                $cabecalho = $colunas[$j]
                $indice = $tabela.ListColumns.Item($cabecalho).Index
                $origem = $linha.Cells.Item(1, $indice)
                $celula = $abaSaida.Cells.Item(2, ($j + 1))
                $abaSaida.Cells.Item(1, ($j + 1)).Value2 = $cabecalho
                $celula.NumberFormat = $origem.NumberFormat
                $valor = $origem.Value2
                if ($valor -is [string] -and $valor -match '^[=+@-]') {
                    $valor = "'" + $valor
                }
                if ($valor -is [string]) {
                    $abaSaida.Cells.Item(2, ($j + 1)).Value2 = [string]$valor
                } elseif ($null -ne $valor) {
                    $abaSaida.Cells.Item(2, ($j + 1)).Value2 = [double]$valor
                }
            }
            $abaSaida.Range('A1:K1').Font.Bold = $true
            $null = $abaSaida.Range('A1:K2').Columns.AutoFit()
            $livroSaida.SaveAs($env:ABACO_DESTINO, 51)
        } catch {
            [Console]::Error.WriteLine($_.Exception.Message + ' (linha ' + $_.InvocationInfo.ScriptLineNumber + ')')
            exit 1
        } finally {
            if ($null -ne $livroSaida) { $livroSaida.Close($false) }
            if ($null -ne $livro) { $livro.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
        }
        """
        pasta = destino.parent
        pasta.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=pasta) as temporaria:
            provisoria = Path(temporaria) / destino.name
            ambiente = os.environ | {
                "ABACO_ORIGEM": str(base),
                "ABACO_DESTINO": str(provisoria),
                "ABACO_MAT": matricula,
                "ABACO_NOME": nome,
                "ABACO_QTDE": quantidade,
            }
            codigo = base64.b64encode(comando.encode("utf-16-le")).decode("ascii")
            processo = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", codigo],
                env=ambiente, capture_output=True, text=True, encoding="utf-8",
                errors="replace", creationflags=subprocess.CREATE_NO_WINDOW, timeout=120,
            )
            if processo.returncode or not provisoria.is_file():
                erro = re.sub(r"<[^>]+>", " ", processo.stderr).strip()
                raise RuntimeError("O Excel não gravou a admissão. " + erro[:250])
            provisoria.rename(destino)

    page.get_by_role("button", name="GESTÃO DE PESSOAS", exact=True).click()
    page.get_by_text("Recursos Humanos", exact=True).wait_for(state="visible")
    url_consulta = urljoin(page.url, "/eturmalina/servlet/hwmdadosfuncionaisservidor")
    resposta = page.goto(url_consulta, referer=page.url, wait_until="domcontentloaded")
    if resposta is not None and resposta.status >= 400:
        raise RuntimeError(f"O Turmalina respondeu HTTP {resposta.status} ao abrir Dados Funcionais do Servidor.")
    campo, consultar = aguardar(formulario, "O campo Matrícula e o botão Consultar não foram identificados.")
    campo.fill(matricula)
    consultar.click()
    quadro, linha, nome_completo, indice_cargo = aguardar(resultado, "A grade não mostrou uma linha identificável para a matrícula 22771 e o vínculo 1.")
    acao = acao_funcional(quadro, linha, indice_cargo)
    if acao is not None:
        acao.click()
    dados = aguardar(ficha, "Abra o segundo ícone depois de Cargo na linha da matrícula 22771 (Eventos funcionais do servidor).", segundos=60 if acao is None else 20)
    numero_na_ficha = re.search(rf"\bservidor\s*:\s*0*{matricula}\b", dados, re.I)
    contrato_na_ficha = re.search(rf"\bcnt\s*:\s*0*{contrato}\b", dados, re.I)
    if not numero_na_ficha or not contrato_na_ficha or norm(nome_completo) not in norm(dados):
        raise RuntimeError("A ficha aberta não confirma o servidor e o contrato demonstrados.")
    if not re.search(r"\badmiss[aã]o\s*:\s*03/11/2026\b", dados, re.I) or not re.search(r"\bjornada\s*:.{0,80}?\b200\s*/\s*40\b", dados, re.I | re.S):
        raise RuntimeError("A ficha não confirma a data de admissão e a jornada do vídeo.")
    nome = " ".join(nome_completo.split()[:2]).upper()
    if norm(nome) != "gileno nascimento":
        raise RuntimeError("O nome lido não coincide com o exemplo.")
    destino = base.parent / "reproducao" / f"{base.stem} - admissao individual {matricula} {datetime.now():%Y%m%d-%H%M%S-%f}.xlsx"
    gravar_admissao(base, destino, nome, "" if original else "19")
