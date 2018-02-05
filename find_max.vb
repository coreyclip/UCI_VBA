Sub GreatyMcGreat():
' finds the Greatest increase and decrease in price change, and the greatest
' total stock volume exchanged. Then this information and the associated stock ticker
' are inputted into a formatted table.
    ' Specify the location of the summary sheet
    ' just make it the first sheet
   With ActiveSheet
        'set row headers
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        ' set column headers
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        ' last row of the percent change column
        LastRow = .Cells(.Rows.Count, 11).End(xlUp).Row
        'create variable for greatest increase and decrease
    End with 
    Dim greatest As Double
    greatest = 0 'greatest increase
    Dim decrease As Double
    decrease = 0 'greatest decrease

    'loop through to find greatest stock price increase
    For i = 2 To LastRow
        Dim val As Double
        val = cells(i, 11).Value
        Dim ticker as String
        ticker = cells(i, 9).Value 
        If val > greatest Then
            greatest = val
            Debug.print "greatest: " & greatest
            Range("P2").Value = ticker 'paste in name of ticker with greatest increase
            Range("Q2").Value = greatest 'paste in value for ticker with greatest increase
            'Range("Q2").NumberFormat = "0.0000%"
        ElseIf val < decrease Then
            decrease = val
            Debug.print "decrease: " & decrease 
            Range("P3").Value = ticker 'paste in name of ticker with greatest decrease
            Range("Q3").Value = decrease 'paste in value for ticker with greatest decrease
            'Range("Q3").NumberFormat = "0.0000%"
        Else

        End If

    Next i
    ' last row of the total stock volume column
    LastRow = Cells(Rows.Count, 12).End(xlUp).Row
    ' create variable that holds greatest volume
    Dim greatest_vol As Double
    greatest_vol = 0
    ' loop through total stock volume to find ticker with the greatest volume
    For i = 2 To LastRow
        Dim volume As Double
        volume = cells(i,12).Value
        ticker = cells(i,9).Value 
        If volume > greatest_vol Then
            greatest_vol = volume
            Debug.print greatest_vol
            Range("Q4").Value = greatest_vol
            Range("P4").Value = ticker
        else 

        End If
    Next i

End Sub