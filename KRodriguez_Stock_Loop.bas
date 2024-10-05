Attribute VB_Name = "Module1"
'Attribute VB_Name = "Stock_Loop"
'@Lang VBA

Sub Stock_loop()

    ' Create variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim ticker_symbol As String
    Dim stock_total As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim Stock_Table_Row As Integer
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestTotalVolumeTicker As String
    Dim greatincrease As Double
    Dim greatdecrease As Double
    Dim greattotal As Double

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

        ' Reset Stock_Table_Row for each worksheet
        Stock_Table_Row = 2
        
        ' Find the last row with data in the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Print headers for Stock Table
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Quarter Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Volume Total"
        
        'Print headers for Greatest Table
        ws.Cells(2, 15).Value = "Greatest Percent Increase"
        ws.Cells(3, 15).Value = "Greatest Percent Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Initialize stock_total
        stock_total = 0
        
        ' Loop through all stock symbols
        For i = 2 To LastRow
        
            ' Check if we are at the last row or if the next stock symbol is different
            If i = LastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the ticker_symbol
                ticker_symbol = ws.Cells(i, 1).Value

                ' Add to the Stock Total
                stock_total = stock_total + ws.Cells(i, 7).Value
                
                ' Calculate the Opening Price
                If ws.Cells(i, 1).Value = ticker_symbol Then
                OpenPrice = ws.Cells(i, 3).Value
                
                End If
                
                ' Calculate the Closing Price
                If ws.Cells(i, 1).Value = ticker_symbol Then
                ClosePrice = ws.Cells(i, 6).Value
                
                End If
                
                ' Calculate the Quarter Change
                quarterly_change = ClosePrice - OpenPrice
                
                ' Calculate the Percent Change
                If OpenPrice <> 0 Then
                    percent_change = Application.WorksheetFunction.Round(((ClosePrice - OpenPrice) / OpenPrice) * 100, 2)
                Else
                    percent_change = 0
                End If
                
                'Calculate the greatest percent increase
                greatincrease = Application.WorksheetFunction.Max(ws.Range("K2:K999999"))
                
                If greatincrease <> 0 Then
                    GreatestIncreaseTicker = ticker_symbol
                Else
                    GreatestIncreaseTicker = 0
                End If
                
                
                'Calculate the greatest percent decrease
                greatdecrease = Application.WorksheetFunction.Min(ws.Range("K2:K999999"))
                
                 If greatdecrease <> 0 Then
                    GreatestDecreaseTicker = ticker_symbol
                Else
                    GreatestDecreaseTicker = 0
                End If
                
                
                'Calculate the greatest total volume
                greattotal = Application.WorksheetFunction.Max(ws.Range("L2:L999999"))
                
                 If greattotal <> 0 Then
                    GreatestVolumeTicker = ticker_symbol
                Else
                    GreatestVolumeTicker = 0
                End If
                
                
                ' Print the ticker symbol in the Stock Table
                ws.Range("I" & Stock_Table_Row).Value = ticker_symbol

                ' Print the Stock Volume to the Stock Table
                ws.Range("L" & Stock_Table_Row).Value = stock_total

                ' Print the Quarter Change to the Stock Table
                ws.Range("J" & Stock_Table_Row).Value = quarterly_change

                ' Print the Percent Change to the Stock Table
                ws.Range("K" & Stock_Table_Row).Value = percent_change
                
                'Print the Greatest Increase
                ws.Cells(2, 16).Value = GreatestIncreaseTicker
                
                 'Print the Greatest Decrease
                ws.Cells(3, 16).Value = GreatestDecreaseTicker
                
                 'Print the Greatest Total Volume
                ws.Cells(4, 16).Value = GreatestVolumeTicker
                
                'Print the Greatest Increase Value
                ws.Cells(2, 17).Value = greatincrease
                
                 'Print the Greatest Decrease Value
                ws.Cells(3, 17).Value = greatdecrease
                
                 'Print the Greatest Total Volume Value
                ws.Cells(4, 17).Value = greattotal

                ' Increment the Stock_Table_Row for the next ticker
                Stock_Table_Row = Stock_Table_Row + 1
                
                ' Reset the stock total for the next ticker
                stock_total = 0
                
            Else
                ' Add to the stock Total
                stock_total = stock_total + ws.Cells(i, 7).Value
            End If
            

        Next i
        
    Next ws
    
End Sub
