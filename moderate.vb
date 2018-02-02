Sub Summarize():

    ' Specify the location of the summary sheet
    ' just make it the first sheet
    Set summary_sheet = Worksheets("A")

    'set column headers for summary tables

    
    summary_sheet.cells(1,8).Value = "Ticker"
    summary_sheet.cells(1,9).Value = "Yearly Change"
    summary_sheet.cells(1,10).Value = "Percent Change"
    summary_sheet.cells(1,11).Value = "Total Stock Volumne"


    ' loop through each worksheet in Worksheets
    For Each ws in Worksheets
        ' create a last row variable
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, 2).End(xlUp).row
        End With
        'create holder for total volume 
        dim total as long 
        total = 0
        dim date_start as Double
        date_start = 2 'first row index of date
        for i = 2 to LastRow
            dim ticker as string
            ticker = ws.cells(i,1).Value 
            if ticker <> ws.cells(i+1,1).value then
                ' calculate out the Yearly change from what the stock opened
                ' the year at to what the closing price was.
                
                ' create a variable for opening price
                Dim opn As Double
                ' create a variable for closing price
                Dim cls As Double
                ' create a variable for change
                Dim chg As Double

                Set date_rng = Sheet1.Range("B" & date_start & ":B" & i)

                opn = WorksheetFunction.Max(date_rng)
                cls = WorksheetFunction.Min(date_rng) 
                
                chg = ws.cells(opn,3).value - ws.cells(cls, 6).value 
                'calculate total stock volume
                nextrow = summary_sheet.Range("I1").SpecialCells(xlCellTypeLastCell).Row
                With summary_sheet
                    .Range("I2:I" & nextrow) = Ticker
                    .Range("J2:J" & nextrow) = chg  
                    .Range("L2:J" & nextrow) = total
                End With
                'incriment start date
                date_start = date_start + i 
            else
                'incriment the total stock volume
                total = total + ws.cells(i,7).value
            End If 
        Next i 
    Next ws     
End Sub
        

        


