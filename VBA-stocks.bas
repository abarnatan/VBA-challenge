Attribute VB_Name = "Module1"
Sub StockSummary()

'Dim ws As Worksheets

For Each ws In Worksheets

Dim Ticker As String

Dim TotalVolume As Double
TotalVolume = 0

Dim YearlyChange As Double

Dim PercentageChange As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim OpeningPrice As Double
OpeningPrice = ws.Cells(2, 3).Value


'Create summary table: add colummn headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Volume"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Double

For i = 2 To lastrow
   
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
    
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    YearlyChange = ws.Cells(i, 6).Value - OpeningPrice
    
    If YearlyChange <= 0 Then
    
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
    Else
    
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
    
    End If
    
    If OpeningPrice = 0 Then
    
    PercentageChange = 0
    
    Else
    
    PercentageChange = YearlyChange / OpeningPrice
    
    End If
    
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    ws.Range("L" & Summary_Table_Row).Value = TotalVolume
    
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
    
    ws.Range("K" & Summary_Table_Row).Value = PercentageChange
    
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    TotalVolume = 0
    
    OpeningPrice = ws.Cells(i + 1, 3).Value
    
    Else
    
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    YearlyChange = ws.Cells(i, 6).Value - OpeningPrice
    
    End If
    
Next i

Next ws

End Sub


