Sub Summarize():

    ' Specify the location of the summary sheet
    ' just make it the first sheet
    Set summary_sheet = Worksheets("A")

    'set column headers for summary tables

    
    summary_sheet.Cells(1, 9).Value = "Ticker"
    summary_sheet.Cells(1, 10).Value = "Yearly Change"
    summary_sheet.Cells(1, 11).Value = "Percent Change"
    summary_sheet.Cells(1, 12).Value = "Total Stock Volumne"


    ' loop through each worksheet in Worksheets
    For Each ws In Worksheets
        ' create a last row variable
        MsgBox(ws.name)
        
        LastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        MsgBox(LastRow)
        'create holder for total volume
        Dim total As Double
        total = 0

        'set initial row position for where opening prices is
        dim row_break as double
        row_break = 2 

        ' create a variable for opening price
        Dim opn As Double
        opn = ws.cells(row_break,3).Value
        'MsgBox(opn) 
        For i = 2 To LastRow
           

            ' create variable for percent change
            Dim percent_change As Double

            'check and see if the next row in the ticker column is the same
            'as the previous ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                MsgBox("New Ticker "& ws.Cells(i + 1, 1).Value )
                'collect closing price 
                cls = ws.Cells(i, 6).Value
                ' create a variable for change
                Dim chg As Double

                chg = opn - cls 
                'calculate percent change
                percent_change = chg / opn
                'set new position of opening price
                'MsgBox(i)
                row_break = row_break + i 
                'create a variable for the ticker
                Dim ticker As String
                ticker = ws.Cells(i, 1).Value

               
                With summary_sheet
                    'last row of the summary table
                     nextrow = .Cells(.Rows.Count, 9).End(xlUp).Row 
                     'MsgBox(nextrow)

                    .Range("I" & nextrow + 1) = ticker
                    .Range("J" & nextrow + 1) = chg
                    .Range("K" & nextrow + 1) = percent_change
                    .Range("J" & nextrow + 1) = total
                End With
                ' set total back to zero 
                total = 0
                'set new opening price 
                opn = ws.Cells(i+1, 3).value 
            Else

                'tabulate total stock volume
                total = total + ws.Cells(i, 7).Value
            end if 
        Next i
    Next ws


End Sub
        

        


