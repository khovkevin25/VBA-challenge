Sub stock_analysis_challenge()
        
    Dim i As Long
    Dim j As Integer
    Dim stockVolume As Double
    Dim lastRow As Long
    Dim stockChange As Double
    Dim percentChange As Double
    Dim summary As Long
        
        
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
        
    j = 0
    stockVolume = 0
    stockChange = 0
    summary = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stockVolume = stockVolume + Cells(i, 7).Value
                
            If stockVolume = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            Else
                If Cells(summary, 3) = 0 Then
                    For find_value = summary To i
                        If Cells(find_value, 3).Value <> 0 Then
                            summary = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                    
                stockChange = (Cells(i, 6) - Cells(summary, 3))
                percentChange = stockChange / Cells(summary, 3)
                summary = i + 1
                    
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = stockChange
                Range("K" & 2 + j).Value = percentChange
                Range("L" & 2 + j).Value = stockVolume
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).NumberFormat = "0.00%"
                    
                Select Case stockChange
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If
                
            stockVolume = 0
            stockChange = 0
            j = j + 1
                
        Else
            stockVolume = stockVolume + Cells(i, 7).Value
        End If
    Next i
        
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastRow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastRow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastRow))
            
    increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
    decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
    volume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastRow)), Range("L2:L" & lastRow), 0)
        
    Range("P2") = Cells(increase + 1, 9)
    Range("P3") = Cells(decrease + 1, 9)
    Range("P4") = Cells(volume + 1, 9)
    Cells.EntireColumn.AutoFit
End Sub
