' ============================================================================
' ThisWorkbook — eventos que cuidam das BDH + copia dos futuros para a exibicao
'
' Cole no editor do VBA (Alt+F11 > ThisWorkbook), substituindo o conteudo
' atual. A estrutura da planilha nao muda: e' so o macro que fica mais esperto.
'
' -------------------------------------------------------------------------
' 1) POR QUE AS BDH SUMIAM
' Os dois eventos assumiam que "o que esta ABAIXO de uma BDH e' o resultado
' derramado dela" -- verdade nas abas onde a BDH fica na linha 4 e a serie
' desce (BBG_Dados, FrangoUS, WEG), e por isso o BeforeSave congelava e o
' Workbook_Open limpava aquele bloco.
'
' Na aba Futuros o layout e' outro: as BDH ficam EMPILHADAS, uma por linha
' (Points=1, uma celula cada). Ali a celula abaixo de uma BDH e' OUTRA BDH,
' entao o BeforeSave congelava as de baixo (formula virava valor) e o
' Workbook_Open apagava o que sobrou -- so a primeira de cada bloco escapava.
'
' Correcao: celula com formula nunca e' derrame, entao os eventos nao mexem
' quando a celula de baixo ja' tem formula. Quem derrama serie segue igual
' (derrame nao tem formula).
'
' -------------------------------------------------------------------------
' 2) DESEMPENHO (o add-in do Bloomberg caia no meio)
' A versao antiga varria "For Each cell In ws.UsedRange": 329 mil celulas,
' quase todas na BBG_Dados (11 mil linhas x 29 colunas). Agora as BDH sao
' localizadas com Find/FindNext, que e' nativo -- sao ~124 celulas, nao 329
' mil. E a limpeza voltou a ser em UM ClearContents por bloco: fazer celula a
' celula dava centenas de milhares de escritas e travava o Excel, derrubando o
' add-in junto.
'
' Os enderecos sao coletados ANTES de alterar qualquer celula: mexer na
' planilha no meio de um Find/FindNext bagunca a proria busca.
'
' -------------------------------------------------------------------------
' 3) COPIA AUTOMATICA PARA A AREA DE EXIBICAO
' Cada bloco da aba Futuros tem duas metades: em cima os valores que a
' dashboard le (e que abrem sem o Bloomberg) e embaixo as formulas vivas.
' Passar uma para a outra era manual. AtualizarExibicaoFuturos faz isso no
' BeforeSave, que roda depois de o Bloomberg ja' ter calculado.
'
' Salvaguarda: se a metade de formulas nao tiver nenhum preco numerico (ex.:
' abriu sem o Bloomberg e tudo virou #N/A), a exibicao NAO e' tocada -- melhor
' manter o ultimo dado bom do que zerar a serie da dashboard.
' ============================================================================

' Coluna do TICKER de cada bloco: B=2 (soja), K=11 (algodao), T=20 (dolar),
' AC=29 (milho). O bloco vai de (ticker-1) a (ticker+3): data, ticker, atual,
' 1 sem., 1 mes.
Private Const COLUNAS_TICKER As String = "2,11,20,29"
Private Const PRIMEIRA_LINHA As Long = 5
Private Const ULTIMA_LINHA As Long = 200


' O add-in do Bloomberg esta REALMENTE respondendo?
'
' Nao basta olhar AddIns(...).Installed: o add-in aparece instalado bem antes de
' registrar as funcoes _xll. Se o Workbook_Open limpar nessa janela, o BDH nao
' tem como derramar a serie de volta -- foi assim que a coluna E da BBG_Dados
' ficou com 1 linha no lugar de 9.737. Aqui a gente pergunta ao proprio add-in.
'
' Detalhe importante: o macro antigo varria 329 mil celulas e demorava, o que
' por acidente dava tempo de o add-in subir antes da limpeza. Ao acelerar a
' busca das BDH, essa protecao acidental caiu -- por isso ela vira explicita.
Private Function BloombergRespondendo() As Boolean
    Dim v As Variant

    On Error Resume Next
    v = Application.Evaluate("BDP(""USDBRL Curncy"",""PX_LAST"")")
    If Err.Number = 0 Then
        If Not IsError(v) Then
            If IsNumeric(v) Then BloombergRespondendo = True
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function


' Espera o add-in ficar pronto. Devolve False se estourar o tempo -- e nesse
' caso o chamador NAO deve limpar nada.
Private Function EsperarBloomberg(Optional segundos As Single = 60) As Boolean
    Dim t0 As Single
    t0 = Timer
    Do
        If BloombergRespondendo() Then
            EsperarBloomberg = True
            Exit Function
        End If
        DoEvents
    Loop While Timer - t0 < segundos
    EsperarBloomberg = False
End Function


' A celula tem algo dentro?
' Comparar .Value com "" estoura "Tipo incompativel" (erro 13) quando a celula
' contem valor de ERRO -- e' o que acontece ao abrir sem o add-in, que enche a
' planilha de #NAME?. Erro conta como conteudo: a celula nao esta vazia.
Private Function TemConteudo(cell As Range) As Boolean
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then
        TemConteudo = True
    Else
        TemConteudo = (v <> "")
    End If
End Function


' Enderecos das celulas cuja formula contem BDH(, nesta aba.
' Devolve os enderecos (nao os Range) de proposito: assim quem chama pode
' alterar a planilha sem interferir na busca.
Private Function EnderecosComBDH(ws As Worksheet) As Collection
    Dim achadas As Collection
    Dim cell As Range
    Dim primeiro As String

    Set achadas = New Collection

    On Error Resume Next
    Set cell = ws.Cells.Find(What:="BDH(", LookIn:=xlFormulas, _
                             LookAt:=xlPart, MatchCase:=False)
    On Error GoTo 0

    If Not cell Is Nothing Then
        primeiro = cell.Address
        Do
            achadas.Add cell.Address
            Set cell = ws.Cells.FindNext(cell)
            If cell Is Nothing Then Exit Do
        Loop While cell.Address <> primeiro
    End If

    Set EnderecosComBDH = achadas
End Function


Private Sub AtualizarExibicaoFuturos()
    Dim ws As Worksheet
    Dim partes() As String
    Dim i As Long, colTk As Long
    Dim rFormula As Long, r As Long
    Dim nLinhas As Long, nValidos As Long
    Dim linha As Long, c As Long
    Dim valorAtual As Variant

    On Error Resume Next
    Set ws = Me.Worksheets("Futuros")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    partes = Split(COLUNAS_TICKER, ",")

    For i = LBound(partes) To UBound(partes)
        colTk = CLng(partes(i))
        ' colTk-1 = data | colTk = ticker | colTk+1 = atual
        ' colTk+2 = 1 sem. | colTk+3 = 1 mes

        ' --- acha onde comeca a metade de formulas deste bloco ---
        rFormula = 0
        For r = PRIMEIRA_LINHA To ULTIMA_LINHA
            If TemConteudo(ws.Cells(r, colTk)) And ws.Cells(r, colTk + 1).HasFormula Then
                rFormula = r
                Exit For
            End If
        Next r

        If rFormula > 0 Then
            ' --- conta as linhas e ve se ha preco de verdade ---
            nLinhas = 0
            nValidos = 0
            r = rFormula
            Do While r <= ULTIMA_LINHA
                If Not TemConteudo(ws.Cells(r, colTk)) Then Exit Do
                nLinhas = nLinhas + 1
                valorAtual = ws.Cells(r, colTk + 1).Value   ' preco atual
                If Not IsError(valorAtual) Then
                    If IsNumeric(valorAtual) And valorAtual <> "" Then
                        nValidos = nValidos + 1
                    End If
                End If
                r = r + 1
            Loop

            ' Salvaguarda: sem nenhum preco numerico (ex.: abriu sem Bloomberg e
            ' virou tudo #N/A), preserva o que ja' esta em cima.
            If nValidos > 0 Then

                ' limpa a exibicao antiga -- o ciclo pode ter encurtado.
                ' para uma linha antes da separacao, para nunca tocar nas formulas
                If rFormula - 2 >= PRIMEIRA_LINHA Then
                    ws.Range(ws.Cells(PRIMEIRA_LINHA, colTk - 1), _
                             ws.Cells(rFormula - 2, colTk + 3)).ClearContents
                End If

                ' copia valores: formulas -> exibicao
                For linha = 0 To nLinhas - 1
                    For c = -1 To 3
                        ws.Cells(PRIMEIRA_LINHA + linha, colTk + c).Value = _
                            ws.Cells(rFormula + linha, colTk + c).Value
                    Next c
                Next linha

            End If
        End If
    Next i
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim ws As Worksheet
    Dim enderecos As Collection
    Dim endereco As Variant
    Dim cell As Range
    Dim outputStart As Range
    Dim lastRow As Long
    Dim outputRng As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Antes de congelar: leva os futuros ja' calculados para a area de exibicao.
    AtualizarExibicaoFuturos

    For Each ws In Me.Worksheets
        Set enderecos = EnderecosComBDH(ws)

        For Each endereco In enderecos
            Set cell = ws.Range(CStr(endereco))
            Set outputStart = cell.Offset(1, 0)

            ' Formula abaixo = BDH empilhada, nao derrame: nao mexe.
            If TemConteudo(outputStart) And Not outputStart.HasFormula Then
                ' Fim do bloco via End(xlDown), que e' nativo.
                If Not TemConteudo(ws.Cells(outputStart.Row + 1, outputStart.Column)) Then
                    lastRow = outputStart.Row
                Else
                    lastRow = outputStart.End(xlDown).Row
                End If

                If lastRow >= outputStart.Row Then
                    Set outputRng = ws.Range(outputStart, ws.Cells(lastRow, outputStart.Column))
                    outputRng.Value = outputRng.Value
                End If
            End If
        Next endereco
    Next ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub Workbook_Open()
    Dim addInAtivo As Boolean
    addInAtivo = False

    Dim addIn As addIn
    For Each addIn In Application.AddIns
        If InStr(1, UCase(addIn.Name), "BLOOMBERG") > 0 And addIn.Installed Then
            addInAtivo = True
            Exit For
        End If
    Next addIn

    ' Instalado nao basta: so limpa depois que as funcoes _xll respondem de
    ' verdade. Sem isso a limpeza pode acontecer na janela em que o add-in ainda
    ' esta subindo, e a serie apagada nao volta.
    If addInAtivo Then
        If Not EsperarBloomberg(60) Then
            ' MsgBox so quando ha' alguem na frente da tela: em automacao
            ' (atualizar_bloomberg.py) uma caixa de dialogo trava o script.
            If Application.Visible Then
                MsgBox "O add-in do Bloomberg nao respondeu a tempo." & vbCrLf & _
                       "Os dados NAO foram limpos, para nao perder as series." & vbCrLf & _
                       "Feche sem salvar, espere o Bloomberg carregar e abra de novo.", _
                       vbExclamation, "Agro.xlsm"
            End If
            Exit Sub
        End If
    End If

    If addInAtivo Then
        Dim ws As Worksheet
        Dim enderecos As Collection
        Dim endereco As Variant
        Dim cell As Range
        Dim outputStart As Range
        Dim lastRow As Long

        Application.ScreenUpdating = False

        For Each ws In Me.Worksheets
            Set enderecos = EnderecosComBDH(ws)

            For Each endereco In enderecos
                Set cell = ws.Range(CStr(endereco))
                Set outputStart = cell.Offset(1, 0)

                ' Formula abaixo = BDH empilhada, nao derrame: nao mexe.
                If Not IsEmpty(outputStart) And Not outputStart.HasFormula Then
                    lastRow = ws.Cells(ws.Rows.Count, outputStart.Column).End(xlUp).Row
                    If lastRow >= outputStart.Row Then
                        ' Um ClearContents por bloco. Celula a celula chegava a
                        ' centenas de milhares de escritas e derrubava o add-in.
                        ws.Range(outputStart, ws.Cells(lastRow, outputStart.Column)).ClearContents
                    End If
                End If
            Next endereco
        Next ws

        Application.ScreenUpdating = True
    End If
End Sub
