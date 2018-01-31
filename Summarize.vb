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

        for i = 2 to LastRow
            dim ticker as string
            ticker = ws.cells(i,1).Value 
            if ticker <> ws.cells(i+1,1).value then
                nextrow = summary_sheet.Range("I1").SpecialCells(xlCellTypeLastCell).Row
                With summary_sheet
                    .Range("I2:I" & nextrow) = Ticker
                    .Range("J2:J" & nextrow) = chg  
                    .Range("L2:J" & nextrow) = total
                End With
            else
                ' calculate out the Yearly change from what the stock opened
                ' the year at to what the closing price was.
                
                ' create a variable for opening price
                Dim opn As Double
                ' create a variable for closing price
                Dim cls As Double
                ' create a varialbe for change
                Dim chg As Double
                opn = ws.Cells(i, 3).Value
                cls = ws.Cells(i, 6).Value
                
                chg = cls - opn
                'calculate total stock volume
            
                total = total + ws.cells(i,7).value

            next i 
    Next ws     


End Sub
        

        


