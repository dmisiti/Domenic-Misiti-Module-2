Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet
Dim lastRow As Long
Dim ticker As String
Dim i As Long
Dim outputRow As Long

Dim openPrice As Double
Dim closePrice As Double
Dim quarterlyChange As Double
Dim percentageChange As Double
Dim stockVolume As Double

Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double
Dim increaseTicker As String
Dim decreaseTicker As String
Dim volumeTicker As String

For Each ws In ThisWorkbook.Worksheets

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    outputRow = 2

    ticker = ws.Cells(2, 1).Value
    openPrice = ws.Cells(2, 3).Value
    closePrice = ws.Cells(2, 6).Value
    stockVolume = ws.Cells(2, 7).Value

    greatestIncrease = -100000
    greatestDecrease = 100000
    greatestVolume = 0

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ticker Then
            quarterlyChange = closePrice - openPrice
            
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice)
            Else
                percentageChange = 0
            End If
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).Value = stockVolume
            
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                increaseTicker = ticker
            End If
            
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                decreaseTicker = ticker
            End If
            
            If stockVolume > greatestVolume Then
                greatestVolume = stockVolume
                volumeTicker = ticker
            End If
            
            outputRow = outputRow + 1
            
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            stockVolume = ws.Cells(i, 7).Value
        
        Else
            closePrice = ws.Cells(i, 6).Value
            stockVolume = stockVolume + ws.Cells(i, 7).Value
        End If
        
    Next i
    
    quarterlyChange = closePrice - openPrice
    
    If openPrice <> 0 Then
        percentageChange = (quarterlyChange / openPrice)
    Else
        percentageChange = 0
    End If
    
    ws.Cells(outputRow, 9).Value = ticker
    ws.Cells(outputRow, 10).Value = quarterlyChange
    ws.Cells(outputRow, 11).Value = percentageChange
    ws.Cells(outputRow, 12).Value = stockVolume
    
    If percentageChange > greatestIncrease Then
        greatestIncrease = percentageChange
        increaseTicker = ticker
    End If
            
    If percentageChange < greatestDecrease Then
        greatestDecrease = percentageChange
        decreaseTicker = ticker
    End If
            
    If stockVolume > greatestVolume Then
        greatestVolume = stockVolume
        volumeTicker = ticker
    End If
    
    ws.Cells(2, 16).Value = increaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 16).Value = decreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 16).Value = volumeTicker
    ws.Cells(4, 17).Value = greatestVolume
    
Next ws
        
End Sub

