'Total Stock Volume
Sub Totstock()

 'Loop Through All the Worksheets
 For Each ws In Worksheets

 'Set our initial variables
 Dim i, j As Long
 Dim totv As Double
 Dim ticker As String
 Dim lastrow As Long
 Dim finalrow As Long
 Dim yearchng As Double
 Dim perchng As Double
 Dim openp As Double
 Dim closep As Double
 Dim op_row As Long
 Dim counter As Long
 Dim maxinc As Double
 Dim maxdec As Double
 Dim maxtotvol As Double
 Dim tickmaxinc As String
 Dim tickmaxdec As String
 Dim tickmaxtotv As String

 'Adds Header Name to appropriate cell
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Value"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"

 'Set Initial for total, counter, and open price to loop through later
 totv = 0
 counter = 2
 op_row = 2
 
 
 'finds the last row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'Loops through each year of stock data
 For i = 2 To lastrow
     
     'Compare Each Ticker to the subsequent ticker
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'gets Ticker when changes
         ticker = ws.Cells(i, 1).Value

         'Calculates Yearly Change and % Change
         openp = ws.Cells(op_row, 3)
         closep = ws.Cells(i, 6)
         yearchng = closep - openp

         'Calculates % Change
         If openp = 0 Then
            perchng = 0
         Else
            perchng = yearchng / openp
         End If

         'Inserts ticker,Total Volume,Yearly Change & % Change values into respective Cells
         ws.Cells(counter, 9).Value = ticker
         ws.Cells(counter, 12).Value = totv + ws.Range("G" & i).Value
         ws.Cells(counter, 10).Value = yearchng
         ws.Cells(counter, 11).Value = perchng
         ws.Cells(counter, 11).NumberFormat = "0.00%"
         
         'Conditional Formatting: + is green; - is red
         If ws.Cells(counter, 10).Value > 0 Then
            ws.Cells(counter, 10).Interior.ColorIndex = 4
         Else
            ws.Cells(counter, 10).Interior.ColorIndex = 3
         End If

         'Adds New Row for display cells for Next Ticker; Sets new open price row and resets total
         counter = counter + 1
         totv = 0
         op_row = i + 1

      
     Else
        'Adds to Total Volume for Each Ticker If Tickers are same
         totv = totv + ws.Cells(i, 7).Value
         
     End If
 Next i

 'sets the initial values for greatest inc, dec, total vol with corresponding tickers to be looped later
 maxinc = ws.Cells(2, 11).Value
 maxdec = ws.Cells(2, 11).Value
 maxtotv = ws.Cells(2, 12).Value
 tickmaxinc = ws.Cells(2, 9).Value
 tickmaxdec = ws.Cells(2, 9).Value
 tickmaxtotv = ws.Cells(2, 9).Value
 
 'finds Last Row Of the cells displayed to be used for finding maxes
 finalrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
 
 'finds greatest value by looping through cells
 For j = 2 To finalrow
     If ws.Cells(j + 1, 11).Value > maxinc Then
        maxinc = ws.Cells(j + 1, 11).Value
        tickmaxinc = ws.Cells(j + 1, 9).Value
     ElseIf ws.Cells(j + 1, 11).Value < maxdec Then
        maxdec = ws.Cells(j + 1, 11).Value
        tickmaxdec = ws.Cells(j + 1, 9).Value
     ElseIf ws.Cells(j + 1, 12).Value > maxtotv Then
        maxtotv = ws.Cells(j + 1, 12).Value
        tickmaxtotv = ws.Cells(j + 1, 9).Value
     End If
 Next j
 
 'Inserts greatest % Inc, greatest % Dec, greatest Total Vol and corresponding ticker into cells
 ws.Cells(2, 16).Value = tickmaxinc
 ws.Cells(3, 16).Value = tickmaxdec
 ws.Cells(4, 16).Value = tickmaxtotv
 ws.Cells(2, 17).Value = maxinc
 ws.Cells(3, 17).Value = maxdec
 ws.Cells(4, 17).Value = maxtotv
 ws.Cells(2, 17).NumberFormat = "0.00%"
 ws.Cells(3, 17).NumberFormat = "0.00%"
 
 'autofits the columns so you can see the text clearly@@
 ws.Columns("I:Q").AutoFit
 
 Next ws
 
End Sub

