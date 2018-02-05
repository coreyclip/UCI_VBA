Sub McFormat():
' Sub formats summary sheets, adds conditional color formatting to yearly change
' adds percentage formatting to percentage change 
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