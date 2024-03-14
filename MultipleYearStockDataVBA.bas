Attribute VB_Name = "Module2"
Sub StockDataAnalysis()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Call ProcessStockData(ws)
    Next ws
End Sub

Sub ProcessStockData(ws As Worksheet)
    'Define my variables
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim rowStart As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    totalVolume = 0
    rowStart = 2

    For i = 2 To lastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerSymbol = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        endPrice = ws.Cells(i, 6).Value
        yearlyChange = endPrice - startPrice
       
        If startPrice <> 0 Then
            percentChange = (yearlyChange / startPrice) * 100
        Else
            percentChange = 0
        End If
        
        ws.Range("I" & rowStart).Value = tickerSymbol
        ws.Range("J" & rowStart).Value = yearlyChange
        ws.Range("K" & rowStart).Value = percentChange
        ws.Range("L" & rowStart).Value = totalVolume
        
    
        If yearlyChange >= 0 Then
            ws.Range("J" & rowStart).Interior.Color = vbGreen
            
        Else
            ws.Range("J" & rowStart).Interior.Color = vbRed
            
        End If
        
        
        If percentChange > greatestIncrease Then
            greatestIncrease = percentChange
            tickerGreatestIncrease = tickerSymbol
            
        ElseIf percentChange < greatestDecrease Then
            greatestDecrease = percentChange
            tickerGreatestDecrease = tickerSymbol
            
        End If
        
        
        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            tickerGreatestVolume = tickerSymbol
            
        End If
        
        
        totalVolume = 0
        
        rowStart = rowStart + 1
       
        startPrice = ws.Cells(i + 1, 3).Value
    Else

        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
    End If
    
Next i

    ws.Range("O2").Value = tickerGreatestIncrease
    ws.Range("P2").Value = greatestIncrease
    ws.Range("O3").Value = tickerGreatestDecrease
    ws.Range("P3").Value = greatestDecrease
    ws.Range("O4").Value = tickerGreatestVolume
    ws.Range("P4").Value = greatestVolume
    
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    Cells(2, 14) = "Greatest % Increase"
    Cells(3, 14) = "Greatest % Decrease"
    Cells(4, 14) = "Greatest Total Volume"


End Sub
