Sub stocks()

    'set dimensions
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim yearlyChange As Double
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim dailyChange As Single
    Dim avgChange As Double
    Dim ws As Worksheet
    
    'create initial for loop
    For Each ws In Worksheets
    
        j = 0
        total = 0
        yearlyChange = 0
        start = 2
        dailyChange = 0
        
        'set title row in each worksheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'get the row number of the last row with data for all worksheets
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'start on first row of data
        For i = 2 To rowCount
        
            'if ticker changes, print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'put results in a variable
                total = total + ws.Cells(i, 7).Value
                
                'if not found
                If total = 0 Then
                
                    'print the results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("l" & 2 + j).Value = 0
                    
                Else
                    
                    If ws.Cells(start, 3) = 0 Then
                    
                        For getValue = start To i
                            
                            If ws.Cells(getValue, 3).Value <> 0 Then
                                
                                start = getValue
                                
                                Exit For
                                
                            End If

                        Next getValue

                    End If
                                
                    yearlyChange = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = yearlyChange / ws.Cells(start, 3)
                            
                    start = i + 1
                            
                    ws.Range("I" & 2 + j) = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j) = yearlyChange
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total
                            
                    'add colors to highlight changes
                    Select Case yearlyChange
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select

                End If
                    
                'reset all counts
                total = 0
                yearlyChange = 0
                j = j + 1
                dailyChange = 0

            Else
                
                'if ticker is still the same, add results
                total = total + ws.Cells(i, 7).Value
                        
            End If
            
        Next i
        
        'take the max and min and incorporate into worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(increaseNumber + 1, 9)
        ws.Range("P3") = ws.Cells(decreaseNumber + 1, 9)
        ws.Range("P4") = ws.Cells(volumeNumber + 1, 9)
        
    Next ws
        
End Sub
