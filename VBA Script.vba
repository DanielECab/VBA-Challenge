'I was receiving having some errors when originally running the code so i used xpert
'to help me add a code to check for errors

Sub ApplyToAllWorksheets()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        Call Multi_year_stock(ws)
        ws.Activate
    Next ws

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred on sheet: " & ws.Name & " - " & Err.Description
End Sub

'Sub script to make sure the code runs on all worksheets in the workbook

Sub ApplyToAllAWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Call Multi_year_stock(ws)
        ws.Activate
    Next ws
End Sub

'Loop script below


Sub Multi_year_stock(ws As Worksheet)

Dim Tickers As String
Dim stockvolume As Double
    stockvolume = 0
Dim sumticker As Integer
    sumticker = 2
Dim open_price As Double
    open_price = Cells(2, 3).Value
Dim close_price As Double
Dim quarterly_change As Double
Dim percent_change As Double
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest %Increase"
    Cells(3, 14).Value = "Greatest %Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        stockvolume = stockvolume + Cells(i, 7).Value
        Range("I" & sumticker).Value = Ticker
        Range("L" & sumticker).Value = stockvolume
        close_price = Cells(i, 6).Value
        quarterly_change = (close_price - open_price)
        Range("J" & sumticker).Value = quarterly_change
    If (open_price = 0) Then
        percent_change = 0
    Else
        percent_change = quarterly_change / open_price
    End If
        Range("k" & sumticker).Value = percent_change
        Range("k" & sumticker).NumberFormat = "0.00%"
        
        sumticker = sumticker + 1
        stockvolume = 0
        open_price = Cells(i + 1, 3)
    Else
        stockvolume = stockvolume + Cells(i, 7).Value
    End If
    
Next i

'Conditional formatting below

Dim cell As Range

    For Each cell In Range("j2:j1501")
        If cell.Value >= 0# Then
            cell.Interior.ColorIndex = 4
        ElseIf cell.Value < 0# Then
            cell.Interior.ColorIndex = 3
        End If
    Next cell
    
    
'Percentage analysis below

Dim percentinc As Double
        percentinc = WorksheetFunction.Max(Range("K:K"))
            Cells(2, "P").Value = percentinc
                Cells(2, "P").NumberFormat = "0.00%"
            
        increase = Application.Match(Cells(2, "P").Value, Range("k:K"), 0)
        Range("O2").Value = Range("I" & increase)
        
        
Dim perecentdec As Double
        percentdec = WorksheetFunction.Min(Range("K:K"))
            Cells(3, "P").Value = percentdec
                Cells(3, "P").NumberFormat = "0.00%"
                
        decrease = Application.Match(Cells(3, "P").Value, Range("K:K"), 0)
        Range("O3").Value = Range("I" & decrease)
        
Dim greatestvolume As Double
        greatestvolume = WorksheetFunction.Max(Range("L:L"))
            Cells(4, "P").Value = greatestvolume
            
        volume = Application.Match(Cells(4, "P").Value, Range("L:L"), 0)
        Range("O4").Value = Range("I" & volume)

End Sub
