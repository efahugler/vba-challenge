Sub VbaofWallStreet()

'Declare every variable
Dim ticker As String

Dim k As Integer

Dim lRow As Long

Dim cp As Double 'closing price

Dim op As Double 'opening price

Dim ychange As Double

Dim total_stock As Double

Dim Per_change As Double



For Each ws In Worksheets

 'Activate each Worksheet
 ws.Activate
 
 
 'Initial conditions
      k = 0 'as a constant of the increasing number of tickers
      ticker = "" 'Number of cells column A
      total_stock = 0 'column G + next row
      op = 0           ' Data given
      ychange = 0     'difference between opening price - closing price
      
 
 'Stablish the headers
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percentage change"
 ws.Range("L1").Value = "Total Stock Volume"

    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
      
For i = 2 To lRow

      
'Ticker column 1
      ticker = Cells(i, 1).Value
      
      'Determine the proper value of the initial stock
      total_stock = 0
      total_stock = total_stock + Cells(i, 7).Value
      
        'Opening price
            op = 0
            If op = 0 Then
            op = Cells(i, 3).Value
            End If
           
'Ticker Conditional
     If Cells(i + 1, 1).Value <> ticker Then
        k = k + 1
        Cells(k + 1, 9) = ticker 'Ticker gets the proper name
           
           
            'Closing price for each ticker
             cp = Cells(i, 6)
             
             'Difference of price
              ychange = 0     'difference between opening price - closing price
              ychange = cp - op
              ychange = Cells(k + 1, 10).Value
              
              'Green & Red
                 If ychange > 0 Then
                   Cells(k + 1, 10).Interior.ColorIndex = 4 'Green
                     ElseIf ychange < 0 Then
                       Cells(k + 1, 10).Interior.ColorIndex = 3 'Red
                           End If
                   
          ' Calculate the percentage change
          Per_change = 0 'Calculate as (yearlychange/opening price) * 100)
           If op > 0 Then
             Per_change = (ychange / op)
               Else
               Per_change = 0
                End If
                  'Storing the % change in the correct format
                               Cells(k + 1, 11).Value = Format(Per_change, "0.00%")

'Total stock on column row 1 column 12
Cells(k + 1, 12).Value = total_stock

'Reset Total Stock Volume to 0 as the iteration of the ticker changes
total_stock = 0
 
 'Reset the opening price to 0 as the ticker changes
opening_price = 0
 
 End If

 Next i
 
 Next ws

End Sub

 'Resources: https://www.tutorialspoint.com/vba/vba_loops.htm'
 'https://www.excelcampus.com/vba/find-last-row-column-cell/#:~:text=To%20find%20the%20last%20used,the%20rows%20in%20the%20worksheet.'
 'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage'
 'https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook' 
