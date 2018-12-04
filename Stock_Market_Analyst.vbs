'   02-VBA_Scripting_Homework_easy
'   Stock_Market_Analyst

 Sub stock_volume()


    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Loop thru each shhet in worksheet
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    For Each ws In Worksheets

        '   set intitial variable for holding the ticker name
        Dim stock_ticker As String
        '   set initial variable for holding total vomume of a particular stock ticker

        Dim total_olume As Long
        total_volume = 0
        
       '    Find last row of data in the column
        Dim last_data_row As Long
        last_data_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        '   Keep track of stock ticker location in summary table row

        Dim ticker_summary_row As Integer
        ticker_summary_row = 2

        '   loop thru all the stocks
        For i = 2 To last_data_row

        '   check to see if we are in the same stock ticker symbol
        '   if not we have reached the end of that stocks data
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                '   set the stock_ticker
                stock_ticker = ws.Cells(i, 1).Value
                '   add to total_volume
                total_value = total_value + ws.Cells(i, 7).Value

                '   print the stock ticker symbol in the summary taable
                ws.Range("I" & ticker_summary_row).Value = stock_ticker
                '   print the Brand Amount to the Summary Table
                ws.Range("J" & ticker_summary_row).Value = total_volume

                '   add 1 to the ticker summary row
                ticker_summary_row = ticker_summary_row + 1
                '   reset the total volume 
                total_volume = 0

        '   if it is take the volume and add it to the total

            Else

                '   add volume to total volume
                total_volume = total_volume + ws.Cells(i, 7).Value

            End If

        Next i

    Next ws 

End Sub 

