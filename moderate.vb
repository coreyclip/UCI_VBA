
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
        
        LastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        MsgBox(LastRow)
        'create holder for total volume
        Dim total As Long
        total = 0

        For i = 2 To LastRow
            ' create a variable for opening price
            Dim opn As Double
            ' create a variable for closing price
            Dim cls As Double
            ' create a variable for change
            Dim chg As Double

            ' create variable for percent change
            Dim percent_change As Double

            'check and see if the next row in the ticker column is the same
            'as the previous ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'create a variable for the ticker
                Dim ticker As String
                ticker = ws.Cells(i, 1).Value
               
                With summary_sheet
                    'last row of the summary table
                     nextrow = .Cells(Rows.Count, 1).End(xlUp).Row
                     MsgBox(nextrow)
                    .Range("I2:I" & nextrow) = ticker
                    .Range("J2:J" & nextrow) = chg
                    .Range("K2:K" & nextrow) = percent_change
                    .Range("L2:J" & nextrow) = total
                End With
                total = 0
            Else
                ' calculate out the Yearly change from what the stock opened
                ' the year at to what the closing price was.
                
                opn = ws.Cells(i, 3).Value
                cls = ws.Cells(i, 6).Value
                
                chg = cls - opn

                'calcualte percent change
                percent_change = chg / opn

                'calculate total stock volume
            
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws


End Sub
        

        


