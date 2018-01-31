Sub percent_change():
    ' calculate percent change over the course of the year
    ' and insert it in as a new column
    ' loop through each worksheet in Worksheets
    For Each ws in Worksheets 
        'create column heading for percent change 
        cells(1,9).value = "<percent change>"
        ' loop through Yearly Change column and opening price to calculate % change
        For i = 2 to ActiveSheet.Cells(ActiveSheet.Rows.Count, 2).End(xlUp).row
            ' bring in opening price as a double 
            dim opn as Double
            opn = cells(i,3).value 

            ' bring in yearly change in absolute terms
            dim yr as double 
            yr = cells(i,8).value 

            ' create variable  change as a percentage, 
            dim chg as percent 

            chg = yr / opn 

            cells(i,9).value = chg 
        next i 
    next ws 

end sub 