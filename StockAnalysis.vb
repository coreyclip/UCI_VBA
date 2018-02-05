Sub Summarize():
    ' loop through each worksheet in Worksheets
    For Each ws In Worksheets
        ws.Range("H1:L50000").Clear

        'set column headers for summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim summary_nextrow As Integer
        summary_nextrow = 2

        ' create a last row variable
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
            Debug.Print "LastRow" & LastRow
        End With

        'create holder for total volume
        Dim total As Double
        total = 0
        Dim date_start As Double
        date_start = 2 'first row index of date

        For i = 2 To LastRow
            Dim ticker As String
            ticker = ws.Cells(i, 1).Value
            If ticker <> ws.Cells(i + 1, 1).Value Then
                ' calculate out the Yearly change from what the stock opened
                ' the year at to what the closing price was.
            

                ' create variable for closing price
                Dim closing As Double
                closing = ws.Cells(i, 6).Value

                ' create a variable for opening price
                Dim opening As Double
                Debug.Print "opening date " & ws.Cells(date_start, 2).Value
                opening = ws.Cells(date_start, 3).Value

                'increment start date
                If date_start < LastRow Then
                    date_start = i + 1
                Else ' we are at the end of the sheet
                    date_start = 2
                End If
                ' create variable for change between opening and closing
                Dim chg As Double
                chg = closing - opening
                
                With ws
                    .Range("I" & summary_nextrow).Value = ticker
                    .Range("J" & summary_nextrow).Value = chg
                    Dim percent_chg As Double
                    If opening = 0 Then 'account for zero division
                        percent_chg = chg
                    Else
                        percent_chg = chg / opening
                    End If
                    .Range("K" & summary_nextrow).Value = percent_chg
                    .Range("L" & summary_nextrow).Value = total
                End With

                ' set total back to zero
                total = 0

                'MsgBox("Next opening date " & date_start)
                'increment next row in summary sheet
                 
                summary_nextrow = summary_nextrow + 1
                'MsgBox("summary_nextrow " & summary_nextrow)
            Else
                'increment the total stock volume
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
    for each ws in Worksheets
        call McFormat
    Next ws
    for each ws in Worksheets
        call GreatyMcGreat
    next ws 
End Sub

Sub McFormat():
    'First Format Yearly Change
    'set up rg as a range, this is a variable to hold our column that will have conditional formating applied to it
    Dim rg As Range
        
    ' we will create three conditional formats, cond1 for greater than 0 aka growth, cond2 for less than 0 aka contraction
    ' and cond3 for equals 0 aka no change at all
    
    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    
    ' set rg to be our yearly change column
    Set rg = Range("J2", Range("J2").End(xlDown))

    'clear any existing conditional formatting, note that in vba we can access elements of an object in a manner similar to python
    rg.FormatConditions.Delete

    'define the rule for each conditional format
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)
    Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, 0)

    'define the format applied for each conditional format, greater than, note how with can be used to have vba sort of autocomplete your code
    ' frankly I think this way of using With is more likely to cause confusion than to just use multi-line editing in an editor,

    With cond1
        .Interior.Color = vbGreen
        .Font.Color = vbBlack
    End With

    ' the above waying of setting up the conditionals is the same as bellow
    cond2.Interior.Color = vbRed
    cond2.Font.Color = vbBlack
    
    
    
    cond3.Interior.Color = vbCyan
    cond3.Font.Color = vbBlack
    ' for reference check: http://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/#ixzz55e5Mijym
    ' set other formatting qualities 
    With ActiveSheet
        .Range("I:O").EntireColumn.AutoFit
        .Range("K:K").EntireColumn.NumberFormat = "0.00%"
        .Range("N:N").ColumnWidth = 5
        .Range("M:M").ColumnWidth = 10
        .Cells(2, 17).NumberFormat = "0.00%"
        .Cells(3, 17).NumberFormat = "0.00%"
        .Cells(4, 17).NumberFormat = "0"
    End With
    
    

End Sub




Sub GreatyMcGreat():
    ' Specify the location of the summary sheet
    ' just make it the first sheet
   With ActiveSheet
        'set row headers
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        ' set column headers
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        ' last row of the percent change column
        LastRow = .Cells(.Rows.Count, 11).End(xlUp).Row
        'create variable for greatest increase and decrease
    End with 
    Dim greatest As Double
    greatest = 0 'greatest increase
    Dim decrease As Double
    decrease = 0 'greatest decrease
    ' create variable that holds greatest volume
    Dim greatest_vol As Double
    greatest_vol = 0

    'loop through to find greatest stock price increase
    For i = 2 To LastRow
        Dim val As Double
        val = Range("K" & i).Value
        Dim ticker as String
        ticker = Range("I" & i).Value 
        Dim volume As Double
        volume = Range("L" & i).Value
        If val > greatest Then
            greatest = val
            Range("P2").Value = ticker 'paste in name of ticker with greatest increase
            Range("Q2").Value = greatest 'paste in value for ticker with greatest increase
            Range("Q2").NumberFormat = "0.0000%"
        ElseIf val < decrease Then
            decrease = val
            Range("P3").Value = ticker 'paste in name of ticker with greatest decrease
            Range("Q3").Value = decrease 'paste in value for ticker with greatest decrease
            Range("Q3").NumberFormat = "0.0000%"
        Else

        End If

    ' loop through total stock volume to find ticker with the greatest volume
        If volume > greatest_vol Then
            greatest_vol = volume
            Debug.print greatest_vol
            Range("Q3").Value = greatest_vol
            Range("P3").Value = ticker
        else 

        End If
    Next i

End Sub

        



