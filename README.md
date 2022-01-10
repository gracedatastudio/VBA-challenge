# VBA-challenge
Sub AnnualStockSummary()

'Set worksheet variable
Dim ws As Worksheet

'Set initial variable for holding the ticker symbol
Dim tickersymbol  As String

'Set initial variable for holding the ticker counter
Dim tickerlabel As Integer
tickerlabel = 2

'Set initial variable for holding the stock opening price
Dim startprice As Double
startprice = 0

'Set initial variable for holding the stock closing price
Dim endprice As Double
endprice = 0

'Set initial variable for holding the yearly change from opening to closing price
Dim yearlychange As Double
yearlychange = 0

'Set initial variable for holding the percentage in change from opening to closing price
Dim percentchange As Double
percentchange = 0

'Set initial variable for holding to total stock volume
Dim totalvolume As Double
totalvolume = 0

For Each ws In Worksheets

'Initial open price
startprice = Cells(2, 3).Value

'Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stocks
For i = 2 To LastRow

    'Check if the ticker symbol is the same, if it is not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'Set ticker symbol
    tickersymbol = Cells(i, 1).Value
   
    'Calculate yearly change
    endprice = Cells(i, 6).Value
    yearlychange = endprice - startprice
    
    'Prevent division by 0
    If startprice <> 0 Then
    percentchange = (yearlychange / startprice) * 100
    
    Else
    
    End If
    
    'Add to the stock volume
    totalvolume = totalvolume + Cells(i, 7).Value
     
    'Print the ticker symbol in summary table
    Range("I" & tickerlabel).Value = tickersymbol
    
    'Print the yearly change
    Range("J" & tickerlabel).Value = yearlychange
    
    'Print the percent change
    Range("K" & tickerlabel).Value = percentchange
    
    'Print the stock volume in the summary table
    Range("L" & tickerlabel).Value = totalvolume
    
    'Add one to the summary table row
    tickerlabel = tickerlabel + 1
    
    'Reset the total stock volume
    totalvolume = 0
    
    'Conditional highlighting
     If Cells(tickerlabel - 1, 10).Value < 0 Then
     Cells(tickerlabel - 1, 10).Interior.ColorIndex = 3
     Else
     Cells(tickerlabel - 1, 10).Interior.ColorIndex = 4
     End If
     
'If the cell immediately following a row is the same ticker
Else

    'Add to the total stock volume
    totalvolume = totalvolume + Cells(i, 7).Value

    
    
End If

Next i

Next ws
    
End Sub

