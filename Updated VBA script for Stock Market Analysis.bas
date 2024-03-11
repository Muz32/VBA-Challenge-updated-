Attribute VB_Name = "Module1"
Sub StockMarketDataAnalysis()


'Assigning variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim PrintAtCell As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim RowCount As Double

'Loop through the three years of worksheet to project output results
For Each ws In Worksheets

'Create columns and put labels
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Assigning values
TotalStockVolume = 0
RowCount = 0
PrintAtCell = 2
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through stock data to find values
For i = 2 To lastRow


    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    
        'To store the stock name
        Ticker = ws.Cells(i, 1).Value

        'To store the number of rows counted until the condition was met
        RowCount = (RowCount + 1) - 1

        'To store values for opening/closing stock prices and total stock volume for each ticker
        ClosingPrice = ws.Cells(i, 6).Value
        OpeningPrice = ws.Cells(i - (RowCount), 3).Value
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        'Formulars for calculations
        YearlyChange = ClosingPrice - OpeningPrice
        PercentChange = YearlyChange / OpeningPrice

        'Printing of output in relevant columns
        ws.Range("I" & PrintAtCell).Value = Ticker
        ws.Range("J" & PrintAtCell).Value = YearlyChange
        ws.Range("L" & PrintAtCell).Value = TotalStockVolume
        ws.Range("K" & PrintAtCell).Value = PercentChange
        ws.Range("K" & PrintAtCell).NumberFormat = "0.00%"

        'Move to the next row to display output
        PrintAtCell = PrintAtCell + 1
    
        'Reset total stock volume and row counter
    
        TotalStockVolume = 0
        RowCount = 0


    Else

        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        RowCount = RowCount + 1


    End If


Next i

'Apply Conditional Formatting to Year change and Percent change columns

lastrowforCF = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

For c = 2 To lastrowforCF

If ws.Cells(c, "J").Value > 0 Then

'Highlights positive change in green
ws.Cells(c, "J").Interior.ColorIndex = 4
ws.Cells(c, "K").Interior.ColorIndex = 4

Else

'Highlights negative change in red
ws.Cells(c, "J").Interior.ColorIndex = 3
ws.Cells(c, "K").Interior.ColorIndex = 3

End If

Next c


Next ws

'To project results for greatest percent increase and decrease, and also greatest total volume of stock.


For Each ws In Worksheets

'Printing row and column headings

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

  
'Assigning variables

Dim TickerValue1 As String
Dim TickerValue2 As String
Dim TickerValue3 As String
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestStockVolume As Double
    
     
'Find the last row in Column K
lastRow1 = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

'Find the last row in Column L
lastRow2 = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row

'Initialize greatest percent increase/decrease and greatest stock volume with the first value

GreatestPercentIncrease = ws.Cells(2, "K").Value
GreatestPercentDecrease = ws.Cells(2, "K").Value
GreatestStockVolume = ws.Cells(2, "L").Value
  
 
'Loop through the remaining rows to find the greatest percent increase and decrease
For i = 2 To lastRow1
        If ws.Cells(i, "K").Value < GreatestPercentDecrease Then
        
            GreatestPercentDecrease = ws.Cells(i, "K").Value
            TickerValue1 = ws.Cells(i, "I").Value
         
        ElseIf ws.Cells(i, "K").Value > GreatestPercentIncrease Then
        
            GreatestPercentIncrease = ws.Cells(i, "K").Value
            TickerValue2 = ws.Cells(i, "I").Value
     
        End If
     
Next i
   
'Loop through the remaining rows to find the greatest total volume value
For j = 2 To lastRow2
       
       If ws.Cells(j, "L").Value > GreatestStockVolume Then
       
            GreatestStockVolume = ws.Cells(j, "L").Value
            TickerValue3 = ws.Cells(j, "I").Value
       
           End If
         
Next j
        
                
'Print Output in relevant rows and columns
 
ws.Cells(2, "Q").Value = GreatestPercentIncrease
ws.Cells(3, "Q").Value = GreatestPercentDecrease
ws.Cells(2, "Q").NumberFormat = "0.00%"
ws.Cells(3, "Q").NumberFormat = "0.00%"
ws.Cells(4, "Q").Value = GreatestStockVolume
ws.Cells(2, "P").Value = TickerValue2
ws.Cells(3, "P").Value = TickerValue1
ws.Cells(4, "P").Value = TickerValue3

ws.Columns("A:P").AutoFit

Next ws


End Sub

