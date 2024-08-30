Sub Stock_Analysis()

Dim ws As Worksheet
Dim wsName As String

ThisWorkbook.Save

For Each ws In ThisWorkbook.Worksheets
    Dim headers1 As Variant
    Dim headers2 As Variant
    Dim headers3 As Variant
    Dim lastrow As Double
    Dim ticker_count As Double
    Dim volume As LongLong
    Dim quarter_open As Double
    Dim quarter_close As Double
    Dim quarter_change As Double
    Dim percent_change As Double
    Dim greatest_increase As Variant
    Dim greatest_decrease As Variant
    Dim greatest_volume As Variant
    
    headers1 = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
    headers2 = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    headers3 = Array("Ticker", "Value")
    ws.Range("I1:L1").Value = headers1
    ws.Range("N2:N4").Value = Application.Transpose(headers2)
    ws.Range("O1:P1").Value = headers3
    
    ticker_count = 2
    volume = 0
    greatest_increase = Array("", 0)
    greatest_decrease = Array("", 0)
    greatest_volume = Array("", 0)
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If volume = 0 Then
            quarter_open = ws.Cells(i, 3).Value
        End If
        
        volume = volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            quarter_close = ws.Cells(i, 6).Value
            quarter_change = quarter_close - quarter_open
            percent_change = (quarter_close - quarter_open) / quarter_open
            
            ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(ticker_count, 10).Value = quarter_change
            ws.Cells(ticker_count, 11).Value = percent_change
            ws.Cells(ticker_count, 12).Value = volume
            
            ws.Cells(ticker_count, 10).NumberFormat = "0.00"
            ws.Cells(ticker_count, 11).NumberFormat = "0.00%"
            
            If percent_change > greatest_increase(1) Then
                greatest_increase(0) = ws.Cells(i, 1)
                greatest_increase(1) = percent_change
            End If
            
            If percent_change < greatest_decrease(1) Then
                greatest_decrease(0) = ws.Cells(i, 1)
                greatest_decrease(1) = percent_change
            End If
            
            If volume > greatest_volume(1) Then
                greatest_volume(0) = ws.Cells(i, 1)
                greatest_volume(1) = volume
            End If
            
            ticker_count = ticker_count + 1
            volume = 0
        
        End If
        
    Next i
    
    Dim rg As Range
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    
    Set rg = ws.Range("J2:J" & lastrow)
    
    Set condition1 = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    With condition1
        .Interior.Color = vbGreen
    End With
    
    Set condition2 = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    With condition2
        .Interior.Color = vbRed
    End With

Next ws


End Sub
