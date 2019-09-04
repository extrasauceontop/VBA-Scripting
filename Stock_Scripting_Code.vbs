Sub swag()

Application.ScreenUpdating = False
   
Dim yearPercent As Double
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    ticks = ActiveSheet.UsedRange.Columns(1).Value
    
    Count = 2
    tickCount = 2
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greates % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Get unique ticker symbols to column I
    '--------------------------------------------------------------------------
    
    For i = 2 To UBound(ticks)
    
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Cells(Count, 9).Value = Cells(i, 1).Value
            Count = Count + 1
        End If
            
    Next i
    
    
    'Get Total Volume, Change, and Percent Change for each ticker
    '----------------------------------------------------------------------------
    
    
    For i = 2 To Count - 1
        totalV = 0
        yearStart = 0
        yearEnd = 0
        For k = tickCount To UBound(ticks)
            
            tickCount = tickCount + 1
            
            'Total Volume Calculation
            If Cells(i, 9).Value = Cells(k, 1).Value Then
                totalV = totalV + Cells(k, 7).Value
            End If
            
            'Get Begining of year value
            If Cells(i, 9).Value = Cells(k, 1).Value And Cells(i, 9).Value <> Cells(k - 1, 1).Value Then
                yearStart = Cells(k, 3).Value
            End If
            
            'Get end of year value
            If Cells(i, 9).Value = Cells(k, 1).Value And Cells(i, 9).Value <> Cells(k + 1, 1).Value Then
                yearEnd = Cells(k, 6).Value
                Exit For
            End If
            
        Next k
        
        'Input total volume
        Cells(i, 12).Value = totalV
        
        'Calculate Year Change
        yearChange = yearEnd - yearStart
        Cells(i, 10).Value = yearChange
        
        'Calculate Percent Change
        If yearStart = 0 Then
            Cells(i, 11).Value = "Undefined"
        Else
            yearPercent = yearChange / yearStart
            Cells(i, 11).Value = yearPercent
        End If
        
    Next i
    
    'Formatting to complete the conditional formatting
    '------------------------------------------------------------------------------------
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range(Cells(2, 10), Cells(UBound(ticks), 10)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("K1").Select
    Selection.FormatConditions.Delete
    Range("A1").Select
    
    'Completing the "Hard" Section
    
    maxiChange = Application.WorksheetFunction.Max(Range("k:k"))
    miniChange = Application.WorksheetFunction.Min(Range("k:k"))
    greatVolume = Application.WorksheetFunction.Max(Range("l:l"))
    
    Range("Q2").Value = maxiChange
    Range("Q3").Value = miniChange
    Range("Q4").Value = greatVolume
    
    Range("Q2").Style = "Percent"
    Range("Q2").NumberFormat = "0.00%"
    
    Range("Q3").Style = "Percent"
    Range("Q3").NumberFormat = "0.00%"
    
    For i = 2 To Count - 1
        If Cells(i, 11).Value = Range("Q2").Value Then
            Range("P2").Value = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value = Range("Q3").Value Then
            Range("P3").Value = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value = Range("Q4").Value Then
            Range("P4").Value = Cells(i, 9).Value
        End If
        
    Next i

Next

starting_ws.Activate

Application.ScreenUpdating = True

End Sub