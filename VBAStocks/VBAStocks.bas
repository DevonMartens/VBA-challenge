Attribute VB_Name = "Module1"
Sub StockMarketTest()
 For Each ws In Worksheets
    ws.Range("I1, O1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest Percent Increase"
    ws.Range("N3").Value = "Greatest Percent Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    Dim Close_Value As Double
    Dim Ticker_Symbol As String
    Dim YOY As Double
    Dim Ticker As String
    
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Summary_Table_Row = 2
    start_row = 2
    TotalStock = 0
    
    For i = 2 To RowCount
        
        Open_Value = ws.Cells(start_row, 3).Value
        TotalStock = TotalStock + ws.Cells(i + 1, 7).Value
               
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Close_Value = ws.Cells(i, 6).Value
            Ticker_Symbol = ws.Cells(i, 1).Value
            YOY = Close_Value - Open_Value
                    If Open_Value = 0 Or Close_Value = 0 Then
                        ws.Range("K" & Summary_Table_Row).Value = 0
                    Else
                        PercentChange = YOY / Open_Value
                    End If
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            ws.Range("J" & Summary_Table_Row).Value = YOY
            ws.Range("L" & Summary_Table_Row).Value = TotalStock
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            Summary_Table_Row = Summary_Table_Row + 1
            start_row = i + 1
            TotalStock = 0
            
               
        End If
    Next i

    ColorLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
  
    For j = 2 To ColorLastRow
            If ws.Cells(j, 10).Value >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(j, 10).Interior.ColorIndex = 3

        End If
    Next j
    
    
    Chart2LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    ws.Cells(2, 16) = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(3, 16) = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(4, 16) = WorksheetFunction.Max(ws.Range("L:L"))
    
    
    RowCount2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
    start_row = 2
    
    For a = 2 To RowCount2
    
              
        If ws.Cells(a, 11).Value = ws.Cells(3, 16).Value Then
            Ticker = ws.Cells(a, 9).Value
            ws.Cells(3, 15).Value = Ticker
            
        ElseIf ws.Cells(a, 11).Value = ws.Cells(2, 16).Value Then
            Ticker = ws.Cells(a, 9).Value
            ws.Cells(2, 15).Value = Ticker
        
        ElseIf ws.Cells(a, 12).Value = ws.Cells(4, 16).Value Then
            Ticker = ws.Cells(a, 9).Value
            ws.Cells(4, 15).Value = Ticker
        Else
        start_row = a + 1
        End If
    Next a
      
      
 Next ws
           

End Sub




