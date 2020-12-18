Sub StockDataScript()

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call SummaryTable1
        Call ConditionalFormatting
        Call SummaryTable2
    Next
    Application.ScreenUpdating = True
End Sub

Sub SummaryTable1()

' Create summary table column headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' Set initial variables
Dim Ticker As String
Dim Yearly_Open As Double
Dim Yearly_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Volume As LongLong

' Keep track of each stock in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Yearly_Open = Range("C2").Value

' Loop through all stock data
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Yearly_Close = Cells(i, 6).Value
        Yearly_Change = Yearly_Close - Yearly_Open

        If Yearly_Open <> 0 Then
            Percent_Change = Yearly_Change / Yearly_Open
        Else
            Percent_Change = 0
        End If
        
        ' Reset yearly open
        Cells(i + 1, 3).Select
        Yearly_Open = ActiveCell.Value
        
        ' Print data in summary table
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Range("L" & Summary_Table_Row).Value = Volume
    
        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0
        Yearly_Open = Range("C" & i).Value
    Else
        Volume = Volume + Cells(i, 7).Value
    End If
    

Next i

' Formatting
Range("I:L").Columns.AutoFit

End Sub
Sub ConditionalFormatting()

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

 ' Conditional formatting
    If Range("J" & Summary_Table_Row).Value > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Summary_Table_Row = Summary_Table_Row + 1
    ElseIf Range("J" & Summary_Table_Row).Value < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        Summary_Table_Row = Summary_Table_Row + 1
    ElseIf Range("J" & Summary_Table_Row).Value = 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = xlNone
    End If

Next i

End Sub
Sub SummaryTable2()

' Create summary table 2 row headers
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

' Create summary table 2 column headers
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

' Greatest % Increase
max_increase = WorksheetFunction.Max(Range("K:K"))
Range("Q2").Value = max_increase
Range("Q2").NumberFormat = "0.00%"

max_ticker = WorksheetFunction.Match(max_increase, Range("K:K"), 0)
Range("P2").Value = Cells(max_ticker, 9)

' Greatest % Decrease
max_decrease = WorksheetFunction.Min(Range("K:K"))
Range("Q3").Value = max_decrease
Range("Q3").NumberFormat = "0.00%"

decrease_ticker = WorksheetFunction.Match(max_decrease, Range("K:K"), 0)
Range("P3").Value = Cells(decrease_ticker, 9)

' Greatest Total Volume
greatest_volume = WorksheetFunction.Max(Range("L:L"))
Range("Q4").Value = greatest_volume

volume_ticker = WorksheetFunction.Match(greatest_volume, Range("L:L"), 0)
Range("P4").Value = Cells(volume_ticker, 9)

' Formatting
Range("O:Q").Columns.AutoFit

End Sub