Attribute VB_Name = "Module1"
Sub Stock_Analysis()
    
    'Declare Variables for each worksheet
    
    For Each ws In Worksheets 'loop through each worksheet
        
       
        Dim ticker_symbol As String ' To hold ticker symbol
        Dim total_vol As Double ' To hold total stock volume
        Dim rowcount As Long  'location tracker for each ticker symbol in the summary table
        Dim year_open As Double  'to hold year opening price
        Dim year_close As Double 'to hold year closing price
        Dim year_change As Double 'to hold the change in price for the year
        Dim lastrow As Long ' to get the total rows to loop through
        
        
        'Assign Variables
        total_vol = 0
        rowcount = 2
        year_open = 0
        year_close = 0
        year_change = 0
        percent_change = 0
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'Create column labels for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"


        'Loop to search through ticker symbols
        For i = 2 To lastrow
            
            'Conditional to get year open price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If
           
            total_vol = total_vol + ws.Cells(i, 7) ' sum up the volume for each row to determine the total stock volume for the year

            'Conditional to determine if the ticker symbol is changing
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value 'print ticker symbol to summary table

               
                ws.Cells(rowcount, 12).Value = total_vol 'print total stock volume to the summary table

                'Grab year end price
                year_close = ws.Cells(i, 6).Value

                'Calculate the price change for the year and print it to the summary table.
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change

                'Conditional to format to highlight positive or negative change.
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Calculate the percent change for the year
                'Conditional for calculating percent change
                If year_open = 0 And year_close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change 'print it to the summary table format as a percentage
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf year_open = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    'Because of this, we only need to evaluate actual price increase by dollar amount and therefore put
                    '"New Stock" as percent change.
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Add 1 to rowcount to move it to the next empty row in the summary table
                rowcount = rowcount + 1

                'Reset total stock volume, year open price, year close price, year change, year percent change
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i
    Next ws
End Sub
