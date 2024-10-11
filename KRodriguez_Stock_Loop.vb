Attribute VB_Name = "Module1"
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
    Dim GreatestIncrease As String
    Dim GreatestDecrease As String
    Dim GreatestTotalVolume As String

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
                
                ' Print the ticker symbol in the Stock Table
                ws.Range("I" & Stock_Table_Row).Value = ticker_symbol

                ' Print the Stock Volume to the Stock Table
                ws.Range("L" & Stock_Table_Row).Value = stock_total

                ' Print the Quarter Change to the Stock Table
                ws.Range("J" & Stock_Table_Row).Value = quarterly_change

                ' Print the Percent Change to the Stock Table
                ws.Range("K" & Stock_Table_Row).Value = percent_change

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
