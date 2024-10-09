Sub stock_analysis_challenge()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        Dim i As Long
        Dim j As Integer
        Dim stockVolume As Double
        Dim lastRow As Long
        Dim stockChange As Double
        Dim percentChange As Double
        Dim summary As Long
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        j = 0
        stockVolume = 0
        stockChange = 0
        summary = 2
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                
                If stockVolume = 0 Then
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                Else
                    If ws.Cells(summary, 3) = 0 Then
                        For find_value = summary To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                summary = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    stockChange = (ws.Cells(i, 6) - ws.Cells(summary, 3))
                    percentChange = stockChange / ws.Cells(summary, 3)
                    summary = i + 1
                    
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = stockChange
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("L" & 2 + j).Value = stockVolume
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    
                    Select Case stockChange
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                
                stockVolume = 0
                stockChange = 0
                j = j + 1
                
            Else
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
            
        increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
        
        ws.Range("P2") = ws.Cells(increase + 1, 9)
        ws.Range("P3") = ws.Cells(decrease + 1, 9)
        ws.Range("P4") = ws.Cells(volume + 1, 9)
    ws.Cells.EntireColumn.AutoFit
    Next ws
End Sub