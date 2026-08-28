' ============================================================================
' ThisWorkbook — versao corrigida dos dois eventos que cuidam das BDH.
'
' Cole no editor do VBA (Alt+F11 > ThisWorkbook), substituindo os dois Subs
' atuais. Nada na planilha muda: e' so o macro que passa a respeitar formulas.
'
' O QUE ESTAVA ERRADO
' Os dois eventos assumiam que "o que esta ABAIXO de uma BDH e' o resultado
' derramado dela" -- verdade nas abas onde a BDH fica na linha 4 e a serie
' desce (BBG_Dados, FrangoUS, WEG), e por isso o BeforeSave congelava e o
' Workbook_Open limpava aquele bloco.
'
' Na aba Futuros o layout e' outro: as BDH ficam EMPILHADAS, uma por linha
' (Points=1, uma celula cada). Ali a celula abaixo de uma BDH e' OUTRA BDH,
' entao o BeforeSave congelava as de baixo (formula virava valor) e o
' Workbook_Open apagava o que sobrou. Resultado: so a primeira BDH de cada
' bloco sobrevivia e o resto sumia -- era preciso reescrever tudo na mao.
'
' A CORRECAO
' Celula com formula nunca e' resultado derramado, e' formula legitima. Entao:
'   1) nao comeca se a celula logo abaixo ja' tem formula;
'   2) o bloco para na primeira celula com formula, em vez de atravessar.
' Quem derrama serie continua funcionando igual (o derrame nao tem formula).
' ============================================================================

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim ws As Worksheet
    Dim cell As Range
    Dim outputStart As Range
    Dim lastRow As Long
    Dim outputRng As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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
