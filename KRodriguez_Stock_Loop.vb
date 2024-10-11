Attribute VB_Name = "Module1"
Sub Stock_loop()

'Create variables
'Set a variable for worksheet loop
Dim ws As Worksheet
Dim LastRow As Long
Dim WorksheetName As String

'Set variable for summary
Dim summarySheet As Worksheet
Dim summaryRow As Long
Dim quarter As String

'Set a variable for holding the counter
Dim i As Long

'Set an intial variable for holding the ticker symbol
Dim ticker_symbol As String

'Set an intial variable for total stock volume
Dim stock_total As Double
stock_total = 0


'set an intial variable for openprice and closingprice
Dim OpenPrice As Double
Dim ClosePrice As Double

' Set an initial variable for holding the quarterly change per ticker symbol
Dim quarterly_change As Double
  quarterly_change = 0

'Set an intial variable for holding the percent change
Dim percent_change As Double
percent_change = 0

            
'Create loop for stock data
  Dim Stock_Table_Row As Integer
  Stock_Table_Row = 2

'Loop through each worksheet in the workbook
  For Each ws In ThisWorkbook.Worksheets
        
 'Find the last row with data in the current worksheet
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
       'Print headers for Stock Table
   ws.Cells(1, 9).Value = "Ticker Symbol"
   ws.Cells(1, 10).Value = "Quarter Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Volume Total"
   
'Output The ticker symbol
'Loop through all stock symbols
'output total stock volume
  For i = 2 To LastRow
  
    ' Check if we are still within the same stock symbol, if it is not...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker_symbol
      ticker_symbol = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      stock_total = stock_total + ws.Cells(i, 7).Value
      
      'calculate the Opening Price
      OpenPrice = Application.WorksheetFunction.Min(ws.Cells(i, 3).Value)
      
      'calculate the Closing Price
      ClosePrice = Application.WorksheetFunction.Max(ws.Cells(i, 6).Value)
      
      'Calculate the Quarter Change
      quarter_change = ClosePrice - OpenPrice
      
      'Output The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
      'Calculate the Percent Change
      percent_change = Application.WorksheetFunction.RoundUp(((ClosePrice - OpenPrice) / OpenPrice) * 100, 2)
      
   
      ' Print the ticker symmbol in the Stock Table
      ws.Range("I" & Stock_Table_Row).Value = ticker_symbol

      ' Print the Stock Volume to the Stock Table
      ws.Range("L" & Stock_Table_Row).Value = stock_total
      
      'Print the Quarter Change to the Stock Table
      ws.Range("J" & Stock_Table_Row).Value = quarter_change
      
      'Print the Percent Change to the Stock Table
      ws.Range("K" & Stock_Table_Row).Value = percent_change

    'Reset the stock total and prices for the next ticker
      stock_total = 0
      OpenPrice = 0
      ClosePrice = 0
                 
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the stock Total
      stock_total = stock_total + ws.Cells(i, 7).Value
      
      
        End If

    Next i
    
Next ws

      

End Sub
