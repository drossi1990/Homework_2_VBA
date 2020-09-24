Sub TickerSummary()
    
    'Begin loop to go through every worksheet in a stock ticker workbook

    Dim sh As Worksheet
    
        For Each sh In Worksheets
        
        
            'Finds last row of worksheet and assigns it to a variable (k) for use in the summary table loop
            Dim rn As Range
            Set rn = sh.UsedRange
            k = rn.Rows.Count + rn.Row - 1
        
    
            'totals trading volume for a particular ticker
            Dim Ticker_Volume As Double
            Ticker_Volume = 0
            
            'defines starting date for a new ticker in the dataset
            Dim Start_Date As Double
            Starting_Date = 0
            
            'defines opening price of a ticker for a given year
            Dim Opening As Double
            Opening = sh.Cells(2, 3).Value
            
            'defines closing price of a ticker for a given year
            Dim Last_Price As Double
            Last_Price = 0
            
            'Tracks next row of worksheet where summary data is to be recorded
            Dim Table_Position As Double
            Table_Position = 2
            
            'track name and volume of ticker with highest trading volume in the sheet
            Dim Greatest_Ticker_Volume As Double
            Greatest_Ticker_Volume = 0
            Dim Greatest_Ticker_Volume_Name As String

            'track name and percent change value of ticker with greatest % gain in value in the sheet
            Dim Greatest_Ticker_PctGain As Double
            Greatest_Ticker_PctGain = 0
            Dim Greatest_Ticker_PctGain_Name As String
            
            'track name and percent change value of ticker with greatest % loss in value in the sheet
            Dim Greatest_Ticker_PctLoss As Double
            Greatest_Ticker_PctLoss = 0
            Dim Greatest_Ticker_PctLoss_Name As String
            
            'Sets header values for summary data for the sheet
            sh.Cells(1, 9).Value = "Ticker"
            sh.Cells(1, 10).Value = "Yearly Change"
            sh.Cells(1, 11).Value = "Percent Change"
            sh.Cells(1, 12).Value = "Total Stock Volume"
            sh.Cells(2, 15).Value = "Greatest % Increase"
            sh.Cells(3, 15).Value = "Greatest % Decrease"
            sh.Cells(4, 15).Value = "Greatest Volume"
            sh.Cells(1, 16).Value = "Ticker"
            sh.Cells(1, 17).Value = "Value"
            
    
    
            'Table summary loop populates summary data for every ticker in a worksheet
            For i = 2 To k
            
                'Ticker_Volume is a counter which tracks which row of the summary table will be populated next by adding volume to a counter which resets on each iteration through the forloop below
                Ticker_Volume = Ticker_Volume + sh.Cells(i, 7).Value
        
                'If the next row's date is less than the current row's date, then the last row for a particular ticker has been reached and summary data can be collected
                If sh.Cells(i + 1, 2).Value < sh.Cells(i, 2).Value Then
                    'Records name of ticker, closing price, and the price change for the summary table
                    Ticker_name = sh.Cells(i, 1).Value
                    Closing = sh.Cells(i, 6).Value
                    Change_Price = Closing - Opening
                    
                    'Finds percent chage, if statement to prevent divisions by zero
                    If Opening <> 0 Then
                        Change_Percentage = (Change_Price / Opening)
                        
                    End If
            
                    'Prints summary table row for current ticker
                    sh.Cells(Table_Position, 9).Value = Ticker_name
                    sh.Cells(Table_Position, 10).Value = Change_Price
                    sh.Cells(Table_Position, 11).Value = Change_Percentage
                    sh.Cells(Table_Position, 11).NumberFormat = "0.00%"
                    sh.Cells(Table_Position, 12).Value = Ticker_Volume
                    
                    'formats positive percent change values, nested loop checks if each value is highest gain in the sheet so far
                    If Change_Percentage > 0 Then
                        
                        sh.Cells(Table_Position, 10).Interior.ColorIndex = 4
                        
                        If Change_Percentage > Greatest_Ticker_PctGain Then
                            Greatest_Ticker_PctGain = Change_Percentage
                            Greatest_Ticker_PctGain_Name = Ticker_name
                        End If
                    
                    'formats negative percent change values, nested loop checks if each value is highest loss in the sheet so far
                    ElseIf Change_Percentage < 0 Then
                        
                        sh.Cells(Table_Position, 10).Interior.ColorIndex = 3
                        
                        If Change_Percentage < Greatest_Ticker_PctLoss Then
                            Greatest_Ticker_PctLoss = Change_Percentage
                            Greatest_Ticker_PctLoss_Name = Ticker_name
                        End If
                    
                    End If
                    
                    'checks if total volume is greatest in the sheet so far
                    If Ticker_Volume > Greatest_Ticker_Volume Then
                        Greatest_Ticker_Volume = Ticker_Volume
                        Greatest_Ticker_Volume_Name = Ticker_name
                    End If
                    
                    
                    'to enable next iteration, resets ticker volume, updates table position, and sets opening to value of next ticker
                Ticker_Volume = 0
                Table_Position = Table_Position + 1
                Opening = sh.Cells(i + 1, 3).Value
            
                End If
    
          
            Next i
        
        'creates summary table for greatest volume, % gain, and % loss on the sheet
        
        sh.Cells(2, 16).Value = Greatest_Ticker_PctGain_Name
        sh.Cells(2, 17).Value = Greatest_Ticker_PctGain
        sh.Cells(2, 17).NumberFormat = "0.00%"
        sh.Cells(3, 16).Value = Greatest_Ticker_PctLoss_Name
        sh.Cells(3, 17).Value = Greatest_Ticker_PctLoss
        sh.Cells(3, 17).NumberFormat = "0.00%"
        sh.Cells(4, 16).Value = Greatest_Ticker_Volume_Name
        sh.Cells(4, 17).Value = Greatest_Ticker_Volume
    
    Next sh
    
End Sub
            
            
 

