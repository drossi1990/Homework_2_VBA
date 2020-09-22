Sub TickerSummary()
    
    'Finds last row of worksheet and assigns it to a variable (k) for use in the summary table loop
    Dim sh As Worksheet
    Dim rn As Range
    Set sh = ThisWorkbook.Sheets("2014")
    Set rn = sh.UsedRange
    k = rn.Rows.Count + rn.Row - 1
    
    'Setting initial values for the ticker summary loop
    Ticker_Volume = 0
    Starting_Date = 0
    Opening = Cells(2, 3).Value
    Last_Price = 0
    Table_Position = 2
    
    
    
    'Table summary loop populates summary data for every ticker in a worksheet
    For i = 2 To k
            
        'Ticker_Volume is a counter which tracks which row of the summary table will be populated next
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        
        'If the next row's date is less than the current row's date, then the last row for a particular ticker has been reached and summary data can be collected
        If Cells(i + 1, 2).Value < Cells(i, 2).Value Then
            'Records name of ticker, closing price, and the price change for the summary table
            Ticker_Name = Cells(i, 1).Value
            Closing = Cells(i, 6).Value
            Change_Price = Closing - Opening
                    
            'Error prevention for division by zero when finding the percentage change of price
            If Opening <> 0 Then
                Change_Percentage = (Change_Price / Opening)
            Else
            
            End If
            
            'Prints summary table row for current ticker
            Cells(Table_Position, 9).Value = Ticker_Name
            Cells(Table_Position, 10).Value = Change_Price
            Cells(Table_Position, 11).Value = Change_Percentage
            Cells(Table_Position, 12).Value = Ticker_Volume
            If Change_Price > 0 Then
                Cells(Table_Position, 10).Interior.ColorIndex = 4
            ElseIf Change_Percentage < 0 Then
                Cells(Table_Position, 10).Interior.ColorIndex = 3
            Else
            
            End If
                
            
            
            'to enable next iteration, resets ticker volume, updates table position, and sets opening to value of next ticker
            Ticker_Volume = 0
            Table_Position = Table_Position + 1
            Opening = Cells(i + 1, 3).Value
            
        End If
        
          
    Next i
            
            
End Sub


