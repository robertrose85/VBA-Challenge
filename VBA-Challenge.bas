Attribute VB_Name = "Module1"
Sub Stock()

Dim ws As Worksheet
Dim flag As Integer
Dim diff_val, open_val, close_val As Double
Dim vol As Variant '16 bytes - no other supported data type (in my version of excel) worked due to value size

red = 3
green = 4

'Loop through each worksheet
For Each ws In Worksheets

    'Set last row of data set
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set starting row
    table_row = 2
    
    'Set Column Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    'Initialize open value
    open_val = ws.Cells(2, 3).Value
    
    'Loop through populated rows.
    For i = 2 To lastrow
    
        'Set column to put unique tickers
        table_col = 9
        
        'Start adding up the volume.
        vol = ws.Cells(i, 7).Value + vol
             
        'If the next column ticker is diff than current column ticker, populate symbol in "Ticker" column. Increment.
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(table_row, table_col).Value = ws.Cells(i, 1).Value
            
            'Assign close value once difference in Ticker is spotted
            close_val = ws.Cells(i, 6).Value
            
            'For testing
            'MsgBox (open_val)
            'MsgBox (close_val)
            
            'Calculate difference between yearly open and close and assign to Yearly Change column
            diff_val = close_val - open_val
            ws.Cells(table_row, 10).Value = diff_val
            
            'Calculate percentage change yearly open and close and assign to Percentage Change column
            percentage_change = (close_val - open_val) / open_val
            ws.Cells(table_row, 11).Value = percentage_change
            ws.Cells(table_row, 11).NumberFormat = "0.00%" ' https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.numberformat
            
            'Change color of cell based on result
            If percentage_change > 0 Then
                ws.Cells(table_row, 10).Interior.ColorIndex = green
            ElseIf percentage_change < 0 Then
                ws.Cells(table_row, 10).Interior.ColorIndex = red
            End If
            
            'Adjust new open value for next Ticker
            open_val = ws.Cells(i + 1, 3).Value
            
            'Assign vol total and reset
            ws.Cells(table_row, 12).Value = vol
            vol = 0
            
            'Increment table_row
            table_row = table_row + 1

        End If
      
    Next i
    
    'Find Min and Max values for Percent Change
    rng_percent = ws.Range("K:K")
    
    ws.Range("Q2") = ws.Application.WorksheetFunction.Max(rng_percent)
    ws.Range("Q3") = ws.Application.WorksheetFunction.Min(rng_percent)
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    ticker_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Grab the corresponding ticker names for percentages
    For j = 2 To 3
    ws.Cells(j, 16).Value = ws.Application.WorksheetFunction.Index(ws.Range("I2:I" & ticker_lastrow), ws.Application.WorksheetFunction.Match(ws.Cells(j, 17).Value, ws.Range("K2:K" & ticker_lastrow), 0))
    Next j
    
    'Greatest Total Volume
    ws.Range("Q4") = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    'Grab ticker names for Total Vol
    ws.Cells(j, 16).Value = ws.Application.WorksheetFunction.Index(ws.Range("I2:I" & ticker_lastrow), ws.Application.WorksheetFunction.Match(ws.Cells(j, 17).Value, ws.Range("L2:L" & ticker_lastrow), 0))
    
    'Auto Fit the columns
    ws.Columns("A:Z").AutoFit
    
    
Next

End Sub


