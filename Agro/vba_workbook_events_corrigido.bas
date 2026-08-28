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
' Correcao: celula com formula nunca e' derrame. Os eventos nao comecam se a
' celula de baixo ja' tem formula, e o bloco para na primeira formula em vez
' de atravessar. Quem derrama serie segue igual (derrame nao tem formula).
'
' -------------------------------------------------------------------------
' 2) COPIA AUTOMATICA PARA A AREA DE EXIBICAO
' Cada bloco da aba Futuros tem duas metades: em cima os valores que a
' dashboard le (e que abrem sem o Bloomberg) e embaixo as formulas vivas.
' Passar uma para a outra era manual. AtualizarExibicaoFuturos faz isso no
' BeforeSave, que roda depois do Bloomberg ja' ter calculado.
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
            If ws.Cells(r, colTk).Value <> "" And ws.Cells(r, colTk + 1).HasFormula Then
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
                If ws.Cells(r, colTk).Value = "" Then Exit Do
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
    Dim cell As Range
    Dim outputStart As Range
    Dim lastRow As Long
    Dim outputRng As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Antes de congelar: leva os futuros ja' calculados para a area de exibicao.
    AtualizarExibicaoFuturos

    For Each ws In Me.Worksheets
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                If InStr(1, UCase(cell.Formula), "BDH(") > 0 Then

                    Set outputStart = cell.Offset(1, 0)

                    ' (1) formula abaixo = BDH empilhada, nao derrame: nao mexe
                    If outputStart.Value <> "" And Not outputStart.HasFormula Then
                        lastRow = outputStart.Row
                        ' (2) para na primeira formula, alem da primeira vazia
                        Do While ws.Cells(lastRow, outputStart.Column).Value <> "" _
                                 And Not ws.Cells(lastRow, outputStart.Column).HasFormula
                            lastRow = lastRow + 1
                        Loop
                        lastRow = lastRow - 1

                        If lastRow >= outputStart.Row Then
                            Set outputRng = ws.Range(outputStart, ws.Cells(lastRow, outputStart.Column))
                            outputRng.Value = outputRng.Value
                        End If
                    End If

                End If
            End If
        Next cell
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

    ' Se Bloomberg estiver presente, limpa os valores estaticos para o BDH atualizar
    If addInAtivo Then
        Dim ws As Worksheet
        Dim cell As Range
        Dim outputStart As Range
        Dim lastRow As Long
        Dim linha As Long

        Application.ScreenUpdating = False

        For Each ws In Me.Worksheets
            For Each cell In ws.UsedRange
                If cell.HasFormula Then
                    If InStr(1, UCase(cell.Formula), "BDH(") > 0 Then

                        Set outputStart = cell.Offset(1, 0)

                        ' (1) formula abaixo = BDH empilhada, nao derrame: nao mexe
                        If Not IsEmpty(outputStart) And Not outputStart.HasFormula Then
                            lastRow = ws.Cells(ws.Rows.Count, outputStart.Column).End(xlUp).Row

                            ' (2) limpa celula a celula e para na primeira formula
                            For linha = outputStart.Row To lastRow
                                If ws.Cells(linha, outputStart.Column).HasFormula Then Exit For
                                ws.Cells(linha, outputStart.Column).ClearContents
                            Next linha
                        End If

                    End If
                End If
            Next cell
        Next ws

        Application.ScreenUpdating = True
    End If
End Sub
