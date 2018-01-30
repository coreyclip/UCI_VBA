Sub yearly_change():
    ' calculate out the Yearly change from what the stock opened
    ' the year at to what the closing price was.
    ' create a last row variable
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, 2).End(xlUp).row
    End With

    ' Set column 8 header to be "yearly change"
    Cells(1, 8) = "yearly change"
    
    For i = 2 To LastRow
        ' create a variable for opening price
        Dim opn As Double
        ' create a variable for closing price
        Dim cls As Double
        ' create a varialbe for change
        Dim chg As Double
        
        opn = Cells(i, 3).Value
        cls = Cells(i, 6).Value
        
        chg = cls - opn
        'Add formatting
        ' FormatConditions.Add(xlCellValue, xlGreater, Formula1, Formula2)
        
        Cells(i, 8).Value = chg
        ' perform conditional formatting regarding weather the change is positive or negative
        Next i
        'set up rg as a range, this is a variable to hold our column that will have conditional formating applied to it
    Dim rg As Range
        
    ' we will create three conditional formats, cond1 for greater than 0 aka growth, cond2 for less than 0 aka contraction
    ' and cond3 for equals 0 aka no change at all
    
    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    
    ' set rg to be our yearly change column
    Set rg = Range("H2", Range("H2").End(xlDown))

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
        .Font.Color = vbWhite
    End With
 
    ' the above waying of setting up the conditionals is the same as b
    cond2.Interior.Color = vbRed
    cond2.Font.Color = vbWhite
    
    
    
    cond3.Interior.Color = vbRed
    cond3.Font.Color = vbWhite
    
    
    

    
    ' for reference check: http://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/#ixzz55e5Mijym


End Sub