Sub VbaofWallStreet()

'Declare every variable
Dim ticker As String

Dim k As Integer

Dim lRow As Long

Dim opening_price As Double

Dim closing_price As Double

Dim yearly_change As Double

Dim percentage_change As Double

Dim total_stock As Double

For Each ws In Worksheets

 'Activate each Worksheet
 ws.Activate
 
 
'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
 
 'Stablish the headers
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percentage change"
 ws.Range("L1").Value = "Total Stock Volume"

    
     'Initial conditions
      k = 0 'as a constant of the increasing number of tickers
      ticker = ""
      opening_price = 0
      percentage_change = 0
      yearly_change = 0
      total_stock_volume = 0
    
For i = 2 To lRow

' Ticker conditional
      ticker = Cells(i, 1).Value
      
        'Opening price
      If opening_price = 0 Then
           opening_price = Cells(i, 3).Value
           End If
           
           'Add up the total Stock volume
           total_stock = total_stock + Cells(i, 7).Value
           
  If Cells(i + 1, 1).Value <> ticker Then
        k = k + 1
        Cells(k + 1, 9) = ticker
           
           
            'Closing price for each ticker
             closing_price = Cells(i, 6)
             
             'Difference of price
              yearly_change = closing_price - opening_price
               Cells(k + 1, 10).Value = yearly_change
              
              'Green & Red
                 If yearly_change > 0 Then
                   Cells(k + 1, 10).Interior.ColorIndex = 4
                     ElseIf yearly_change < 0 Then
                       Cells(k + 1, 10).Interior.ColorIndex = 3
                           End If
                   
          ' Calculate the percentage change
           If opening_price = 0 Then
             percentage_change = 0
               Else
               percentage_change = (yearly_change / opening_price)
                End If
                
                'Storing the % change in the correct format
                Cells(k + 1, 11).Value = Format(percentage_change, "0.00%")

'Total stock on column row 1 column 12
Cells(k + 1, 12).Value = total_stock

'Reset Total Stock Volume to 0 as the iteration of the ticker change
total_stock = 0
 
 'Reset the opening price to 0 as the ticker changeCells
opening_price = 0
 
 End If

 Next i
 
 Next ws

End Sub
 'I help myself with the following links https://www.excelcampus.com/vba/find-last-row-column-cell/#:~:text=To%20find%20the%20last%20used,the%20rows%20in%20the%20worksheet.' 
 'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage'
 'https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook'
 