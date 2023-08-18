Sub StockData()

Dim Worksheet As Worksheet
    For Each Worksheet In ThisWorkbook.Worksheets
        Worksheet.Select

Dim Ticker As String
Dim YearlyChange As Double
Dim SummaryTableRow As Integer
Dim PercentChange As Double
Dim StockVolume As Variant
Dim GreatestPercentIncrease As Variant
Dim GreatestPercentDecrease As Variant
Dim TotalStock As Variant

    YearlyChange = 0
    StockVolume = 0
    SummaryTableRow = 2
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        YearlyChange = YearlyChange + (Cells(i, 6).Value - Cells(i, 3).Value)
        StockVolume = StockVolume + Cells(i, 7).Value
        Range("I" & SummaryTableRow).Value = Ticker
        Range("J" & SummaryTableRow).Value = YearlyChange
        PercentChange = YearlyChange / (Cells(i, 6).Value - YearlyChange)
        Range("K" & SummaryTableRow).Value = FormatPercent(PercentChange)
            If YearlyChange < 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            ElseIf YearlyChange >= 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            End If
        Range("L" & SummaryTableRow).Value = StockVolume
        SummaryTableRow = SummaryTableRow + 1
        YearlyChange = 0
        StockVolume = 0
    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        YearlyChange = YearlyChange + (Cells(i + 1, 3).Value - Cells(i, 6).Value) + (Cells(i, 6).Value - Cells(i, 3).Value)
        StockVolume = StockVolume + Cells(i, 7).Value
    End If

Next i
    GreatestPercentIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    GreatestPercentDecrease = Application.WorksheetFunction.Min(Range("K:K"))
    TotalStock = Application.WorksheetFunction.Max(Range("L:L"))

For i = 2 To lastrow
    If Cells(i, 11).Value = GreatestPercentIncrease Then
        Range("P2").Value = Cells(i, 9).Value
        Range("Q2").Value = FormatPercent(Cells(i, 11).Value)
    ElseIf Cells(i, 11).Value = GreatestPercentDecrease Then
        Range("P3").Value = Cells(i, 9).Value
        Range("Q3").Value = FormatPercent(Cells(i, 11).Value)
    ElseIf Cells(i, 12).Value = TotalStock Then
        Range("P4").Value = Cells(i, 9).Value
        Range("Q4").Value = Cells(i, 12).Value
    End If

Next i

Next Worksheet

End Sub
