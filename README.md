<!-- # Givens_VBA_challenge -->
# VBA Homework - The VBA of Wall Street

## Objective

Analyze three years (2014-2016) of stock market data from .xlsx document utilizing recently obtained VBA scripting knowledge.


### Steps

1. Opened Github to create a new repository called 'Givens_VBA_challenge` for this project.

2. Navigated in terminal to the homework VU Documents folder and cloned the new respository.

3. While still inside my local respository, typed command <mkdir VBAstocks> to create a directory for to house any VBA files that will hold the scripts for each analysis.

4. While working again in terminal a couple days after first configuring the new repository I rain into "refusing to merge unrelated histories" error while attempting to push updates. Final solution was to create a new repository and git add all existing working files. 

5. Finally pushed the above changes to GitHub repo.


## Instructions for VBA Scpripting 



* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

  ' Declare all variables to loop through
    Dim ws As Worksheet
    Dim ticker As String
    Dim vol As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer

    'Ran into overflow error, found below statement on StackOverflow
    On Error Resume Next

    'Start For loop to run through each worksheet one at a time
      For Each ws In ThisWorkbook.Worksheets
        'Set all headers in row 1 to name the data we are about to summarize
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    'Setup where the summary table is going to be inputting in row 2
      Summary_Table_Row = 2

    'Begin For loop to start sorting through data
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Run through columns to record all the values from each previsouly named variables
            ticker = ws.Cells(i, 1).Value
            vol = ws.Cells(i, 7).Value

            year_open = ws.Cells(i, 3).Value
            year_close = ws.Cells(i, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close

            'Collect and insert the values into the summary table
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

             vol = 0
        
           Else: vol = vol + ws.Cells(i, 7)
        End If

        

'The end of the for loop
    Next i
    
ws.Columns("K").NumberFormat = "0.00%" -->

  You should also have conditional formatting that will highlight positive change in green and negative change in red.

    'Declare the format for columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    'For loop to fill in each cell in above declare range that highlights positive change in green and negative change in red
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g


'Continue and run loop/module on the next worksheet
Next ws -->