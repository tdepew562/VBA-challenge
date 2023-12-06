Sub StocksFINAL()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ' Headers
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

        ' Variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim Volume As Double
        Dim OpenStock As Double
        Dim CloseStock As Double
        Dim lastrow As Long
        Dim Summary_Table_Row As Long

        ' Finding the last row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initializing variables
        Volume = 0
        Summary_Table_Row = 2

        ' Loop through data
        For a = 2 To lastrow
            If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
                ' Calculations
                Ticker = ws.Cells(a, 1).Value
                Volume = Volume + ws.Cells(a, 7).Value
                CloseStock = ws.Cells(a, 6).Value

                If OpenStock = 0 Then
                    YearlyChange = 0
                    PercentChange = 0
                Else
                    YearlyChange = CloseStock - OpenStock
                    PercentChange = YearlyChange / OpenStock
                End If

                ' Output to summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Volume
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                ' Reset variables for the next ticker
                Volume = 0
                Summary_Table_Row = Summary_Table_Row + 1
            ElseIf ws.Cells(a - 1, 1).Value <> ws.Cells(a, 1).Value Then
                ' Opening stock for the next ticker
                OpenStock = ws.Cells(a, 3)
            Else
                ' Accumulating volume for the same ticker
                Volume = Volume + ws.Cells(a, 7).Value
            End If
        Next a

        ' Conditional formatting for Yearly Change
        For YearlyFormat = 2 To lastrow
            If ws.Range("J" & YearlyFormat).Value > 0 Then
                ws.Range("J" & YearlyFormat).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & YearlyFormat).Value < 0 Then
                ws.Range("J" & YearlyFormat).Interior.ColorIndex = 3
            End If
        Next YearlyFormat

        ' Conditional formatting for Percent Change
        For PercentFormat = 2 To lastrow
            If ws.Range("K" & PercentFormat).Value > 0 Then
                ws.Range("K" & PercentFormat).Interior.ColorIndex = 4
            ElseIf ws.Range("K" & PercentFormat).Value < 0 Then
                ws.Range("K" & PercentFormat).Interior.ColorIndex = 3
            End If
        Next PercentFormat

        ' Headers for the second table
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")

        ' Headers for the third table
        ws.Range("O2:O4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")

        ' Variables for the third table
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double

        ' Initializing variables for the third table
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0

        ' Loop for finding greatest values
        For increaseIndex = 2 To lastrow
            If ws.Cells(increaseIndex, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(increaseIndex, 11).Value
                ws.Range("Q2").Value = GreatestIncrease
                ws.Range("Q2").Style = "Percent"
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = ws.Cells(increaseIndex, 9).Value
            End If
        Next increaseIndex

        For decreaseIndex = 2 To lastrow
            If ws.Cells(decreaseIndex, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(decreaseIndex, 11).Value
                ws.Range("Q3").Value = GreatestDecrease
                ws.Range("Q3").Style = "Percent"
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Cells(decreaseIndex, 9).Value
            End If
        Next decreaseIndex

        For volumeIndex = 2 To lastrow
            If ws.Cells(volumeIndex, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(volumeIndex, 12).Value
                ws.Range("Q4").Value = GreatestVolume
                ws.Range("P4").Value = ws.Cells(volumeIndex, 9).Value
            End If
        Next volumeIndex
    Next ws
End Sub


