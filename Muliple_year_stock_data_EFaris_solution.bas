Attribute VB_Name = "Module1"
Sub Stock_Analysis()

' Declare Variables
Dim ticker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentChange As String
Dim totalVolume As Double
Dim Summary_Table_Row As Double

Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double
Dim increaseTicker As String
Dim decreaseTicker As String
Dim volumeTicker As String



' Loop though all stocks in each worksheet
For Each ws In Worksheets

    '  Find last row in each worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create column headers for each output variable
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Initialize variables
    openingPrice = ws.Cells(2, 3).Value
    Summary_Table_Row = 2
    totalVolume = 0
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    
    ' Loop through rows in worksheet and create summary table
        For i = 2 To lastRow
            ' Check if in the same ticker and define values
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    closingPrice = ws.Cells(i, 6).Value
                    yearlyChange = closingPrice - openingPrice
                    percentChange = Format((yearlyChange / openingPrice), "0.00%")
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
            ' Print values in Summary
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = yearlyChange
                ws.Cells(Summary_Table_Row, 11).Value = percentChange
                ws.Cells(Summary_Table_Row, 12).Value = totalVolume
            ' Update variables for new stock
                openingPrice = ws.Cells(i + 1, 3).Value
                Summary_Table_Row = Summary_Table_Row + 1
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
            End If
            
                
        Next i
        
    
    ' Use conditional formatting to differentiate positive and negative yearly change
    lastSRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    For i = 2 To lastSRow
        If ws.Cells(i, 10) > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
    Next i
    
    ' Find the greatest percent increase, decrease, and total volume
    
    ' Create Table
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
       ' Loop through the summary rows to determine values
        For i = 2 To lastSRow
            If CDbl(ws.Cells(i, 11).Value) > greatestIncrease Then
                greatestIncrease = CDbl(ws.Cells(i, 11).Value)
                increaseTicker = ws.Cells(i, 9).Value
            End If
            
            If CDbl(ws.Cells(i, 11).Value) < greatestDecrease Then
                greatestDecrease = CDbl(ws.Cells(i, 11).Value)
                decreaseTicker = ws.Cells(i, 9).Value
            End If
        
            If ws.Cells(i, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(i, 12).Value
                volumeTicker = ws.Cells(i, 9).Value
            End If
        
        Next i

        ' Display values in worksheet
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Display values as percent
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        
    ' Autofit to display data
    ws.Columns("I:P").AutoFit
    
Next ws
        
End Sub
