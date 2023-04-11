ption Explicit
Sub Stocks()
    ' For each loop to run script across all sheets in workbook
   
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
         
        ' define constants write to summary table 
        Const FIRST_DATA_ROW As Integer = 2
        Const INPUT_VOL_COL As Integer = 7
        Const TICKER_COL As Integer = 1
        Const OPEN_COL As Integer = 3
        Const CLOSE_COL As Integer = 6
        ' define variables write to summary table
        Dim Ticker As String
        Dim Total_Stock_Volume As LongLong
        Dim LastRow As Long
        Dim First_Open As Double
        Dim Last_Close As Double
        Dim Yearly_Change As Double
        Dim Summary_Table_Row As Long
        Dim Input_Row As Long
        Dim Yearly_Change_Frac As Double
        ' Set initial values for variables
        First_Open = Cells(FIRST_DATA_ROW, OPEN_COL).Value
        Total_Stock_Volume = 0
        Summary_Table_Row = FIRST_DATA_ROW
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' For loop ticker/changes/volume
        For Input_Row = FIRST_DATA_ROW To LastRow
            Ticker = Cells(Input_Row, TICKER_COL).Value
            Total_Stock_Volume = Total_Stock_Volume + Cells(Input_Row, INPUT_VOL_COL).Value
            If Cells(Input_Row + 1, TICKER_COL).Value <> Ticker Then
            ' inputs
            Last_Close = Cells(Input_Row, CLOSE_COL).Value
            ' calculations
            Yearly_Change = Last_Close - First_Open
            Yearly_Change_Frac = Yearly_Change / First_Open
            ' outputs
            Range("i" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = FormatPercent(Yearly_Change_Frac)
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            ' Nested loops for conditional formatting
            Dim FormatNeg As Long
            For FormatNeg = 1 To 1
                If Range("j" & Summary_Table_Row).Value < 0 Then
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
                Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            Next FormatNeg
            Dim FormatPos As Long
            For FormatPos = 1 To 1
                If Range("j" & Summary_Table_Row).Value > 0 Then
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
                Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
            Next FormatPos
            ' set up for next row in stock table
            First_Open = Cells(Input_Row + 1, OPEN_COL).Value
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Stock_Volume = 0
            End If
        Next Input_Row

        ' print header to summary table
        Range("i1").Value = "Ticker"
        Range("j1").Value = "Yearly Change"
        Range("k1").Value = "Percent Change"
        Range("l1").Value = "Total Stock Volume"

        ' Define variables/constants to calculate stocks greatest increase/decrease/total volume
        Const FIRST_SUMMARY_ROW As Integer = 2
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As LongLong
        Dim TickerInc As String
        Dim TickerDec As String
        Dim TickerVol As String
        Dim LastSummaryRow As Long
        Dim j As Long
        GreatestIncrease = -1
        GreatestDecrease = 0
        GreatestVolume = -1
        LastSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row
        ' for loop 
        For j = FIRST_SUMMARY_ROW To LastSummaryRow
            If Cells(j, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(j, 11).Value
            TickerInc = Cells(j, 9).Value
            End If
            If Cells(j, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(j, 11).Value
            TickerDec = Cells(j, 9).Value
            End If
            If Cells(j, 12) > GreatestVolume Then
            GreatestVolume = Cells(j, 12).Value
            TickerVol = Cells(j, 9).Value
            End If
        Next j

        ' print header/values to second summary table
        Range("p1").Value = "Ticker"
        Range("q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("p2").Value = TickerInc
        Range("p3").Value = TickerDec
        Range("p4").Value = TickerVol
        Range("q2").Value = FormatPercent(GreatestIncrease)
        Range("q3").Value = FormatPercent(GreatestDecrease)
        Range("q4").Value = GreatestVolume
        
    Next ws
    End Sub