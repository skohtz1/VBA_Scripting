Sub test2()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim lCol As Long
    Dim ticker As String
    Dim stockTotal As Double
    

    'Location for each total and ticker
    Dim Summary_table_row As Integer

            

    For Each ws In Worksheets
        ws.Range("I1") = "ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Difference"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        Summary_table_row = 2
        stockTotal = 0
        
        Dim open1 As Double
        Dim close1 As Double
        Dim yeardiff As Double
        Dim percDiff As Double
        
        'Find the last non-blank cell in column A(1)
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Find the last non-blank cell in row 1
        lCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        'Set the first open value
        open1 = ws.Range("C2")
        For i = 2 To lRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Check if the ticker name has changed from the previous cell, if it has then:
                'Set the ticker name

                ticker = ws.Cells(i, 1).Value
                'Set the closing market value
                close1 = ws.Cells(i, 6).Value
                'Set the difference from opening the year to ending the year
                yeardiff = close1 - open1
                'Make sure yeardiff and open1 is not 0, this helps avoid a division by 0 error
                If yeardiff <> 0 and open1 > 0 Then
                
                    percDiff = Round((yeardiff / open1), 4)
                    
                Else
                    percDiff = 0
                    
                End If
                
                'Setting the cells equal to the percentage format with the correct format
                ws.Range("K" & Summary_table_row).Value = percDiff
                    
                ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                
                ws.Range("J" & Summary_table_row).Value = yeardiff

                'Conditional Formatting for the percent change
                
                If yeardiff <= 0 Then
                
                    ws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
                    
                End If

                'Adding up the stock totals for the specific named ticker
        
                stockTotal = stockTotal + ws.Cells(i, 7).Value
        
                ws.Range("I" & Summary_table_row).Value = ticker
        
                ws.Range("L" & Summary_table_row).Value = stockTotal
                
                
                
                'Moving on to the next row in our new table when there is a new ticker name
                Summary_table_row = Summary_table_row + 1
                'setting the stock total = 0 for new ticker
                stockTotal = 0
                
                open1 = ws.Cells(i + 1, 3)
        
            Else
                stockTotal = stockTotal + ws.Cells(i, 7).Value
                
            End If
        
        Next i

        'Find the highest percentage increase
        ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("k:k"))
        ws.Range("Q2").NumberFormat = "0.00%"
        'Find  the highest percentage decrease
        ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("k:k"))
        ws.Range("Q3").NumberFormat = "0.00%"
        'Find the highest total volume
        ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("l:l"))
        
        For j = 2 To lRow
            If ws.Range("K" & j) = Application.WorksheetFunction.Max(ws.Range("k:k")) Then
                Dim tick3 As String
                tick3 = ws.Range("I" & j)
                ws.Range("P2") = tick3
                
            ElseIf ws.Range("K" & j) = Application.WorksheetFunction.Min(ws.Range("k:k")) Then
                Dim tick2 As String
                tick2 = ws.Range("I" & j)
                ws.Range("P3") = tick2
                
            End If
            
            If ws.Range("L" & j) = Application.WorksheetFunction.Max(ws.Range("l:l")) Then
                Dim tick1 As String
                tick1 = ws.Range("I" & j)
                ws.Range("P4") = tick1
                
                
            End If
            
        Next j
            
            
    Next ws

End Sub