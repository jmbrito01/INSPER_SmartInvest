Attribute VB_Name = "Module1"
'Module1 File
'Written by João Marcelo Brito
'10/05/2015
'=====================================================================================================================
Public bTimer As Boolean

Function cycles() As Integer
    Dim c As Integer
    c = 1
    Do While (Worksheets("MainSheet").Cells(c, 1) <> vbNullString)
        c = c + 1
    Loop
    cycles = c - 3
End Function

Function GetAroon() As Single
    Dim c As Integer
    Dim max As Single
    Dim maxIdx As Integer
    c = 3
    Do While (Worksheets("MainSheet").Cells(c, 1) <> vbNullString)
        If (Worksheets("MainSheet").Cells(c, 1) >= max) Then
           max = Worksheets("MainSheet").Cells(c, 1)
           maxIdx = cycles - c
        End If
        c = c + 1
    Loop
    c = c - 3
    GetAroon = 100 * ((c - maxIdx + 2) / c)
End Function

Function GetAvarage() As Single
    Dim c As Integer
    Dim sum As Single
    c = 3
    If (cycles = 0) Then
        GetAvarage = 0
    Else
        Do While (Worksheets("MainSheet").Cells(c, 1) <> vbNullString)
        sum = sum + Worksheets("MainSheet").Cells(c, 1)
        c = c + 1
        Loop
        GetAvarage = sum / (c - 3)
    End If
End Function

Function GetVolatility() As Single
    Dim c As Integer
    Dim avg As Single
    Dim sum As Single
    Dim dev As Single
    c = 3
    avg = GetAvarage
    If (avg = 0) Then
        GetVolatility = 0
    Else
        Do While (Worksheets("MainSheet").Cells(c, 1) <> vbNullString)
            dev = dev + (Worksheets("MainSheet").Cells(c, 1) - avg) ^ 2
            c = c + 1
        Loop
        GetVolatility = (Sqr(dev / (c - 3)) / avg) * 100
    End If
End Function

Private Sub UpdateStock()
    Dim connectstring As String
    Dim w As Worksheet
    Dim b As Workbook
    Dim count As Integer
    Dim Aroon As Single
    Dim i As Integer
    
    Set w = ActiveWorkbook.Worksheets.Add()
    connectstring = "URL;http://br.advfn.com/bolsa-de-valores/bovespa/" + MainFrm.cbStock.Text + "/cotacao"
    With w.QueryTables.Add(Connection:=connectstring, Destination:=w.Range("A1"))
        .Name = "tmpWorksheet"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "4"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    DoEvents
    MainFrm.lblMax.Caption = "Máximo: R$" + CStr(w.Cells(2, 5))
    MainFrm.lblMin.Caption = "Mínimo: R$" + CStr(w.Cells(2, 6))
    MainFrm.lblMedia.Caption = "Média: R$" + CStr(GetAvarage)
    MainFrm.lblPrice.Caption = "Preço da Ação: R$" + CStr(w.Cells(2, 4))
    Worksheets("MainSheet").Cells(cycles + 3, 1) = w.Cells(2, 4)
    Application.DisplayAlerts = False
    w.Delete
    Application.DisplayAlerts = True
    Worksheets("MainSheet").Cells(cycles + 2, 2) = GetAroon / 100
    MainFrm.lblVolatile.Caption = "Volatilidade: " + CStr(GetVolatility) + "%"
    If (GetVolatility > 2) Then
        MainFrm.lblAlert.Caption = "Alerta Volatilidade: Alta"
    Else
        MainFrm.lblAlert.Caption = "Alerta Volatilidade: Baixa"
    End If
    MainFrm.lblAroon.Caption = "Índice Aroon: " + CStr(GetAroon) + "%"
    If (GetAroon >= 70) Then
        MainFrm.lblRecom.Caption = "Recomendação: Comprar ação!"
        Worksheets("MainSheet").Cells(cycles + 2, 3) = "Comprar!"
    Else
        Worksheets("MainSheet").Cells(cycles + 2, 3) = "Vender!"
        MainFrm.lblRecom.Caption = "Recomendação: Vender ação!"
    End If
    
    If (Worksheets("MainSheet").ChartObjects(1).Visible = True) Then
        Worksheets("MainSheet").ChartObjects(1).Activate
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.SetSourceData Source:=Range(Worksheets("MainSheet").Cells(3, 1), Worksheets("MainSheet").Cells(cycles + 2, 1))
    End If
    
    If (cycles > 1) Then
        If (Worksheets("MainSheet").Cells(cycles + 2, 1) = Worksheets("MainSheet").Cells(cycles + 1, 1)) Then
            Worksheets("MainSheet").Cells(cycles + 2, 4) = "SEM MUDANÇA"
        ElseIf (Worksheets("MainSheet").Cells(cycles + 1, 3) = "Comprar!") Then
            If (Worksheets("MainSheet").Cells(cycles + 2, 1) > Worksheets("MainSheet").Cells(cycles + 1, 1)) Then
                Worksheets("MainSheet").Cells(cycles + 2, 4) = "ACERTOU!"
            Else
                Worksheets("MainSheet").Cells(cycles + 2, 4) = "ERROU!"
            End If
        ElseIf (Worksheets("MainSheet").Cells(cycles + 1, 3) = "Vender!") Then
            If (Worksheets("MainSheet").Cells(cycles + 2, 1) < Worksheets("MainSheet").Cells(cycles + 1, 1)) Then
                Worksheets("MainSheet").Cells(cycles + 2, 4) = "ACERTOU!"
            Else
                Worksheets("MainSheet").Cells(cycles + 2, 4) = "ERROU!"
            End If
        End If
    End If
    
    count = 0
    For i = 3 To cycles + 3
        If (Worksheets("MainSheet").Cells(i, 4) = "ACERTOU!") Then
            count = count + 1
        End If
    Next i
    MainFrm.lblStatus.Caption = "Acertos: " + CStr(count) + "/" + CStr(cycles) + "(" + CStr((count / (cycles)) * 100) + "%)"
End Sub

Sub StockTimer()
    If (bTimer) Then
        UpdateStock
        Application.OnTime Now + TimeSerial(0, 0, CInt(MainFrm.txtDelay.Text)), "StockTimer"
    End If
End Sub
'End of file
'=====================================================================================================================
