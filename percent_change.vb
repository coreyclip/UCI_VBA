Sub percent_change():
    ' calculate percent change over the course of the year
    ' and insert it in as a new column
    ' loop through each worksheet in Worksheets
    For Each ws in Worksheets 

        ' loop through Yearly Change column and opening price to calculate % change
        For i = 2 to ActiveSheet.Cells(ActiveSheet.Rows.Count, 2).End(xlUp).row
            ' create variable  change as a percentage, 
            dim chg as percent 
            ' bring in opening price as a double 
            dim opn as Double
            opn = cells(i,) 
            ' bring in yearly change in absolute terms 

        next i 


end sub 