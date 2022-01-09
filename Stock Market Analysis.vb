Sub Market_Analysis()
'For loop to loop through sheets
Dim Sheetcount As Integer
Sheetcount = ActiveWorkbook.Worksheets.Count
For j = 1 To Sheetcount
Worksheets(j).Range("I1").Value = "Ticker"
Worksheets(j).Range("J1").Value = "Yearly change"
Worksheets(j).Range("K1").Value = "Percent change"
Worksheets(j).Range("L1").Value = "Total Stock Volume"

' Declare and Initialize variables
Dim Tickr As String
Dim YearlyChange As Double
Dim StockOpen As Double
Dim StockClose as Double
Dim Previous as Long
Previous = 2
Dim PercChange As Double
PercChange = 0
Dim SumVolume As Double
SumVolume = 0
Dim DisplayRow As Integer
DisplayRow = 2

'For loop to loop through rows for a given stock
Dim NumRows As Long
NumRows = Cells(Rows.Count,"A").End(xlUp).Row

For i = 2 To NumRows
SumVolume = SumVolume + Cells(i, 7).Value
'If statement to see if stock has changed
    If Worksheets(j).Cells(i, 1).Value <> Worksheets(j).Cells(i + 1, 1).Value Then
       
        ' Calculate and Display Results
        StockOpen = Worksheets(j).Range("C" & Previous).Value
        StockClose = Worksheets(j).Range("F" & i)
        YearlyChange = StockClose - StockOpen

        Tickr = Worksheets(j).Cells(i, 1).Value
        Worksheets(j).Cells(DisplayRow, 9) = Tickr
        Worksheets(j).Cells(DisplayRow, 10) = YearlyChange
        If YearlyChange >= 0 Then
            Worksheets(j).Cells(DisplayRow, 10).Interior.ColorIndex = 4
        Else
            Worksheets(j).Cells(DisplayRow, 10).Interior.ColorIndex = 3
        End If
        If StockOpen = 0 Then
            PercChange = 0
        Else 
            PercChange = YearlyChange / StockOpen
        End If
        Worksheets(j).Cells(DisplayRow, 11) = PercChange
        Worksheets(j).Cells(DisplayRow, 12) = SumVolume

        ' Reset variables for next stock
        SumVolume = 0
        DisplayRow = DisplayRow + 1
        Previous = i + 1
    End If
Next i
Next j
End Sub