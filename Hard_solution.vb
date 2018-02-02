Sub GreatyMcGreat():
     Specify the location of the summary sheet
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
    LastRow = summary_sheet.Cells(ws.Rows.Count, 11).End(xlUp).Row
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
        elseif val < decrease then 
            decrease = val 
        else 

        end if 
    next i 
    ' loop through 
    for i = 2 to LastRow ' last row of total stock volume should be the same 
        
        

Sub End 