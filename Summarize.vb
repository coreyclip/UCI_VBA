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
        ' gather ticker

        dim ticker as string
        ticker = ws.cells(2,1).Value 


        ' calculate out the Yearly change from what the stock opened
        ' the year at to what the closing price was.
        
        ' create a variable for opening price
        Dim opn As Double
        ' create a variable for closing price
        Dim cls As Double
        ' create a varialbe for change
        Dim chg As Double


        ' create a last row variable
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, 2).End(xlUp).row
        End With

        opn = ws.Cells(2, 3).Value
        cls = ws.Cells(LastRow, 6).Value
        
        chg = cls - opn
         With summary_sheet
            LastRow = .Cells(.Rows.Count, 2).End(xlUp).row
            .Range("I2:I" & LastRow) = Ticker
            .Range("J2:J" & LastRow) = chg  
        End With
        
        'TODO set chg to correct row, make sure you're not writing over
        ' the previous worksheet's data 
        
        
        
        
    Next ws     


End Sub