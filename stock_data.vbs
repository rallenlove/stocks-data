Sub totalVolume()
    
    For Each ws In Worksheets
    
        ws.Select
        
        ' Clear results cells (mainly for testing)
        Columns("I:N").ClearContents
        Range("P2:Q4").ClearContents
        
        ' Set labels
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("Q2:Q4").Value = 0
        
        ' Format % columns
        Columns("K").NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Delete conditional formatting rules
        Columns("J:J").FormatConditions.Delete

        ' Conditional formatting for "Yearly Change" column
        Columns("J:J").Select
        Columns("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Columns("J:J").FormatConditions(Columns("J:J").FormatConditions.Count).SetFirstPriority
        With Columns("J:J").FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
        End With
        Columns("J:J").FormatConditions(1).StopIfTrue = False
        Columns("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Columns("J:J").FormatConditions(Columns("J:J").FormatConditions.Count).SetFirstPriority
        With Columns("J:J").FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Columns("J:J").FormatConditions(1).StopIfTrue = False

        ' I like filters and frozen top row
        Columns("A:G").Select
        Selection.AutoFilter
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        Dim ticker
        Dim previous_ticker
        Dim total_position
        Dim open_price
        Dim close_price
        Dim increase
        Dim pct_change

        ' Keep track of position to place ticker values/aggregate values
        total_position = 1
        
        ' MsgBox (Range("A2").End(xlDown))
        
        ' Iterate over data cells
        ' +1 because otherwise we miss the last ticker (doesn't trigger <> condition)
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            previous_ticker = ticker
            ticker = Range("A" + CStr(i)).Value
            
            ' Perform operations every time we reach a new ticker
            If ticker <> previous_ticker Then

                ' Operations for close price stop after row 1, the header
                If total_position > 1 Then
                
                    close_price = Range("F" + CStr(i - 1)).Value
                    ' Range("M" + CStr(total_position)).Value = open_price
                    ' Range("N" + CStr(total_position)).Value = close_price
                    increase = close_price - open_price
                    Range("J" + CStr(total_position)).Value = increase
                    
                    ' Prevents divide by 0
                    If open_price > 0 Then
                        Range("K" + CStr(total_position)).Value = increase / open_price
                    End If
                    
                End If
                
                total_position = total_position + 1
                Range("I" + CStr(total_position)).Value = ticker
                open_price = Range("C" + CStr(i)).Value
    
            End If
            
            total_cell = "L" + CStr(total_position)
            
            Range(total_cell).Value = Range(total_cell).Value + Range("G" + CStr(i)).Value
        
        Next i

        ' iterate over totals to get overall pct changes and "greatest volume"
        For j = 2 To total_position
            
            ticker = Range("I" + CStr(j)).Value
            pct_change = Range("K" + CStr(j)).Value

            If pct_change > Range("Q2").Value Then
                Range("P2").Value = ticker
                Range("Q2").Value = pct_change
            End If
            
            If pct_change < Range("Q3").Value Then
                Range("P3").Value = ticker
                Range("Q3").Value = pct_change
            End If

            If Range("L" + CStr(j)).Value > Range("Q4").Value Then
                Range("P4").Value = ticker
                Range("Q4").Value = Range("L" + CStr(j)).Value
            End If

        Next j
        
        ' Fit new columns
        Columns("J:L").EntireColumn.AutoFit
        Columns("O:Q").EntireColumn.AutoFit
        
        ' O:Q wasn't visible without scrolling
        Columns("M:N").ColumnWidth = 5
        
        ' Reset view
        Range("A2").Select

    Next

Worksheets(1).Select

End Sub
