Sub GreatyMcGreat():
    ' Specify the location of the summary sheet
    ' just make it the first sheet
    Set summary_sheet = Worksheets("A")
    'set row headers
    summary_sheet.range("O2").value = "Greatest % Increase"    
    summary_sheet.range("O3").value = "Greatest % Decrease"
    summary_sheet.range("O4").value = "Greatest Total Volume"
    ' set column headers 
    summary_sheet.range("P1").value = "Ticker"
    summary_sheet.range("Q1").value = "Value"
    ' last row of the percent change column
    LastRow = summary_sheet.Cells(summary_sheet.Rows.Count, 11).End(xlUp).Row
    'create variable for greatest increase and decrease
    dim greatest as double 
    greatest = 0 'greatest increase
    dim decrease as double
    decrease = 0 'greatest decrease

    'loop through to find greatest stock price increase
    for i = 2 to LastRow 
        dim val as double 
        val = cells(i,11).value
        
        if val > greatest then
            greatest = val
            range("P2").value = cells(i,9).value 'paste in name of ticker with greatest increase 
            range("Q2").value = greatest 'paste in value for ticker with greatest increase
            range("Q2").NumberFormat = "0.0000%"  
        elseif val < decrease then 
            decrease = val 
            range("P3").value = cells(i,9).value 'paste in name of ticker with greatest decrease
            range("Q3").value = greatest 'paste in value for ticker with greatest decrease
            range("Q3").NumberFormat = "0.0000%"  
        else 

        end if 

    next i 
    ' create variable that holds greatest volume
    dim greatest_vol As double
    greatest_vol = 0
    ' loop through total stock volume to find ticker with the greatest volume
    for i = 2 to LastRow ' last row of total stock volume should be the same 
        dim volume as double 
        volume = cells(i, 12).value 
         
        if volume > greatest_vol then 
            greatest_vol = volume 
            range("P3").value = greatest_vol
        else 
        end if 
    
    next i 
        

End Sub 