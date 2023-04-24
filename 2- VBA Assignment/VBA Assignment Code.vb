Sub StockData()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate
Application.ScreenUpdating = False

    'Row Count
    Dim rc As Long
    'Row iterator
    Dim r As Long
    'Unique Ticker counter
    Dim TickerCnt As Integer
    'Total of all Tickers
    Dim Total As Double
    'Iterator for Second Table
    Dim i As Integer

    rc = 0
    r = 0
    TickerCnt = 0
    Total = 0
    i = 0
    
    'Count of #rows
    rc = Cells(Rows.Count, 1).End(xlUp).Row
        
    'DATA NEEDS TO BE SORTED BY TICKER AND DATE FOR THE LOGIC TO WORK:

        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
        ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range("A2:A" & rc), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & rc), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
        With ActiveWorkbook.ActiveSheet.Sort
            .SetRange Range("A1:G" & rc)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    ActiveWindow.DisplayGridlines = False

    'Declaring size of arrays thats hold Unique Tickers, TotalbyTicker volume, open and close Values
    ReDim b(rc) As String
    ReDim t(rc) As Double
    ReDim Op(rc) As Double
    ReDim Cl(rc) As Double

    ReDim GI(0, 1) As Variant
    ReDim GD(0, 1) As Variant
    ReDim GT(0, 1) As Variant

    'Initialize the unique Ticker count to 0
    TickerCnt = 0
    'Register the first Ticker Name
    b(TickerCnt) = Cells(2, 1)
    'Open Value of first Ticker
    Op(TickerCnt) = Cells(2, 3).Value
    'TotalByTicker to first Ticker volume
    t(TickerCnt) = Cells(2, 7).Value

    
    
    'Loop through each row
    For r = 2 To rc - 1

    ' If a row is the same as next row add to the TotalByTicker
        If Cells(r, 1) = Cells(r + 1, 1) Then
            t(TickerCnt) = t(TickerCnt) + Cells(r + 1, 7).Value
        Else
        'If a row is now different from next row then
        'Register previous ticker closing value
            Cl(TickerCnt) = Cells(r, 6).Value
        'increase the Ticker counter
        'Register the new Ticker Name
        'Register the new opening value
        'initialize the TotalByTicker to next Ticker Volume
            TickerCnt = TickerCnt + 1
            b(TickerCnt) = Cells(r + 1, 1)
            Op(TickerCnt) = Cells(r + 1, 3).Value
            t(TickerCnt) = Cells(r + 1, 7).Value
        End If
    Next r


    'Assigning labels to the tables
    Total = 0
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"

    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
        
        
    'Populate the 2nd table

            For i = 0 To TickerCnt
            
                Cells(i + 2, 9).Value = b(i)
                Cells(i + 2, 10).Value = Cl(i) - Op(i)
                Cells(i + 2, 11).Value = (Cl(i) / Op(i)) - 1
                    If Cells(i + 2, 11).Value <= 0 Then
                    Cells(i + 2, 10).Interior.Color = RGB(255, 0, 0)
                    Cells(i + 2, 11).Interior.Color = RGB(255, 0, 0)
                    Else
                    Cells(i + 2, 10).Interior.Color = RGB(0, 255, 0)
                    Cells(i + 2, 11).Interior.Color = RGB(0, 255, 0)
                    End If
                Cells(i + 2, 12).Value = t(i)
                Total = Total + t(i)
                
            Next i
        
    'Get the values for the 3rd table

    GI(0, 0) = Range("K2").Value
    GI(0, 1) = Range("K2").Offset(0, -2).Value
    GD(0, 0) = Range("K2").Value
    GD(0, 1) = Range("K2").Offset(0, -2).Value
    GT(0, 0) = Range("L2").Value
    GT(0, 1) = Range("L2").Offset(0, -3).Value

            For i = 2 To TickerCnt
            
                If GI(0, 0) < Range("K" & i + 1) Then
                GI(0, 0) = Range("K" & i + 1).Value
                GI(0, 1) = Range("K" & i + 1).Offset(0, -2).Value
                End If
                
                If GD(0, 0) > Range("K" & i + 1) Then
                GD(0, 0) = Range("K" & i + 1).Value
                GD(0, 1) = Range("K" & i + 1).Offset(0, -2).Value
                End If
                
                If GT(0, 0) < Range("L" & i + 1) Then
                GT(0, 0) = Range("L" & i + 1).Value
                GT(0, 1) = Range("L" & i + 1).Offset(0, -3).Value
                End If
                
            Next i


    'Populate te 3rd table

                Cells(2, 16) = GI(0, 0)
                Cells(2, 15) = GI(0, 1)
                
                Cells(3, 16) = GD(0, 0)
                Cells(3, 15) = GD(0, 1)
                
                Cells(4, 16) = GT(0, 0)
                Cells(4, 15) = GT(0, 1)

    'FORMAT THE SHEET(recorded Macro - no fancy logic here):

    Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        Range("I1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With

        Range("N1:P4").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        Range("A1:G1,I1:L1,N2:N4,O1:P1").Select
        Range("O1").Activate
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With

        Range("K2:K" & rc).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"

        Range("P2:P3").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    
        Range("L2:L" & rc).Select
        Selection.Style = "Currency"
        Selection.NumberFormat = "$0,0"

        Range("P4").Select
        Selection.Style = "Currency"
        Selection.NumberFormat = "$0,0"
    
        Range("A1").Select

Next ws

End Sub



