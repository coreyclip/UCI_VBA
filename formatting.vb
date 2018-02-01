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
    ' Format Percent Change to be an actual percentage
    Set rg = Range("K2", Range("K2").End(xlDown))
    for each cell in rg
        cell.NumberFormat = "0.0000%" 
    next cell 
End Sub