Sub stockData()


Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    Dim ticker As String
    Dim yearStart As Double
    Dim yearEnd As Double
    Dim percent As Double
    Dim total As LongLong
    Dim lastRow As Long
    Dim counter As Long
    
    Dim incValue As Double
    Dim decValue As Double
    Dim totalValue As LongLong
    
    incValue = 0
    decValue = 0
    totalValue = 0
    
    counter = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    yearStart = Cells(2, 3).Value
    
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    
    
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearEnd = Cells(i, 6).Value
            yearly = yearEnd - yearStart
            If yearStart = 0 Then
                percent = 0
            Else
                percent = yearly / yearStart
            End If
            
            total = total + Cells(i, 7).Value
            Cells(counter, 9).Value = Cells(i, 1).Value
            Cells(counter, 10).Value = yearly
                If Cells(counter, 10).Value < 0 Then
                    Cells(counter, 10).Interior.ColorIndex = 3
                ElseIf Cells(counter, 10).Value > 0 Then
                    Cells(counter, 10).Interior.ColorIndex = 4
                End If
            Cells(counter, 11).Value = FormatPercent(percent)
            Cells(counter, 12).Value = total
            
        
                
            
            
            counter = counter + 1
            yearStart = Cells(i + 1, 3).Value
            percent = 0
            total = 0
            
            
            
        Else
            total = total + Cells(i, 7).Value
        
        End If
        
        

        Next i
    
    For i = 2 To lastRow
    If Cells(i, 11).Value > incValue Then
        incValue = Cells(i, 11).Value
        Range("p2").Value = Cells(i, 9).Value
        Range("q2").Value = FormatPercent(incValue)
    ElseIf Cells(i, 11).Value < decValue Then
        decValue = Cells(i, 11).Value
        Range("p3").Value = Cells(i, 9).Value
        Range("q3").Value = FormatPercent(decValue)
    ElseIf Cells(i, 12).Value > totalValue Then
        totalValue = Cells(i, 12).Value
        Range("p4").Value = Cells(i, 9).Value
        Range("q4").Value = totalValue
    
    End If
    Next i
        
    
    Next WS



End Sub
