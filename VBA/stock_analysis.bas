Sub stocks()

'This program will scan through an excel spreadsheet with stock data for multiple years with each year on its own sheet and determine the 
'yearly change, percent change and total stock value for each stock. It will also determine which stock had the greatest % increase, % decrease and total volume per year.
'For the code to run properly, the tickers must be sorted so the tickers are grouped together but the dates do not need to be in order.

    'This loops through each sheet, running the analysis
    For Each ws In Worksheets
        
        Dim count As Double
        Dim last_day As Double
        Dim first_day As Double
        Dim dif As Double
        Dim reset As Boolean
        Dim vol As Double
        Dim map As Double
        Dim min As Double
        Dim mav As Double
        Dim ind_map As Double
        Dim ind_min As Double
        Dim ind_mav As Double
        
        'Initializes the count at 2 to start displaying the analysis of the stocks in the 2nd row under the headers and volume at 0 to sum it up
        count = 2
        vol = 0
        
        'Determine the final row and creating the headers for the new columns
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            
            'This conditional will find the index for the opening and closing days of each ticker. When the for loop reaches a new ticker, this will be reset
            If i = 2 Then
            
                last_day = i
                first_day = i
                
            ElseIf reset = True Then
                
                last_day = i
                first_day = i
                reset = False
                
            ElseIf ws.Cells(i, 2).Value > last_day Then
            
                last_day = i
                
            ElseIf ws.Cells(i, 2).Value < first_day Then
            
                first_day = i
            
            End If
            
            'This conditional tests if the ticker value after each row is different. If it is, it will calculate the yearly change and % change
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Displays the ticker in the first new column, making one row for each unique ticker
                ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
                
                'Calculates the difference between the value for the opening on the first day and the closing on the final day, assigns the value in the table and formats the numbers
                dif = ws.Cells(last_day, 6).Value - ws.Cells(first_day, 3).Value
                ws.Cells(count, 10).Value = dif
                ws.Cells(count, 10).NumberFormat = "0.00"
                
                'Formats the cells so that they are green if the change is positive and red if the change is negative
                If dif > 0 Then
                
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                    
                ElseIf dif < 0 Then
                
                    ws.Cells(count, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Calculates the percent change from the opening on the first day and assigns it to the corresponding cell.
                'Formats it to display as a percentage
                ws.Cells(count, 11).Value = dif / ws.Cells(first_day, 3).Value
                ws.Cells(count, 11).NumberFormat = "0.00%"
                
                'Adds the final value to the volume and assigns it to the cell
                vol = vol + ws.Cells(i, 7).Value
                ws.Cells(count, 12).Value = vol
                ws.Cells(count, 12).NumberFormat = "0"
                
                'Resets the values and ensures that it moves to the next row for the next unique ticker
                count = count + 1
                vol = 0
                reset = True
                
            Else
                
                'Running total of the volume for the ticker
                vol = vol + ws.Cells(i, 7).Value
            
            End If
            
        Next i

       'Creating the new column and row labels 
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Loops through the unique tickers
        For i = 2 To count

            'Initializes the starting values for greatest % increase, % decrease and total volume with their respective indices
            If i = 2 Then
            
                map = ws.Cells(i, 11).Value
                min = ws.Cells(i, 11).Value
                mav = ws.Cells(i, 12).Value
                ind_map = i
                ind_mav = i
                ind_min = i
            
            'Tests each ticker to see if they had the greatest % increase, % decrease or total volume and if they do it saves their index
            ElseIf ws.Cells(i, 11).Value > map And ws.Cells(i, 12).Value > mav Then
            
                map = ws.Cells(i, 11).Value
                mav = ws.Cells(i, 12).Value
                ind_map = i
                ind_mav = i
                
            ElseIf ws.Cells(i, 11).Value < min And ws.Cells(i, 12).Value > mav Then
            
                min = ws.Cells(i, 11).Value
                mav = ws.Cells(i, 12).Value
                ind_mav = i
                ind_min = i
                
            ElseIf ws.Cells(i, 11).Value > map Then
            
                map = ws.Cells(i, 11).Value
                ind_map = i
                
            ElseIf ws.Cells(i, 11).Value < min Then
            
                min = ws.Cells(i, 11).Value
                ind_min = i
            
            ElseIf ws.Cells(i, 12).Value > mav Then
            
                mav = ws.Cells(i, 12).Value
                ind_mav = i

            End If
        
        Next i
        
        'Once it is done testing the tickers, it inputs the values in the cells as well as labeling which ticker corresponds to which value
        ws.Cells(2, 16).Value = Cells(ind_map, 9).Value
        ws.Cells(2, 17).Value = map
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = Cells(ind_min, 9).Value
        ws.Cells(3, 17).Value = min
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = Cells(ind_mav, 9).Value
        ws.Cells(4, 17).Value = mav
    
    Next ws
    
End Sub
