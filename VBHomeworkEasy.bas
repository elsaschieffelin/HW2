Attribute VB_Name = "Module1"
Sub TotalVolumeStocks():

For Each ws In Worksheets

'Create Ticker and Total Stock Volume Columns
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Total Stock Volume"
    ws.Columns("K:K").AutoFit

'Set Variables for rows and total volume
    Dim StockRow As Long
    Dim TotalVolume As Double

'Set StockRow equal to 2 so we can start printing tickers in a clean list and TotalVolume equal to 0
    StockRow = 2
    TotalVolume = 0
'Find the Last Row
    Dim LastRow As Long

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'identify the change in stock and add up the total stock volume for each stock in 2016
    Dim i As Long
    For i = 2 To LastRow
    'Begin to tally up the total stock volume
       TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    'Identify where the stock ticker changes
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Record ticker
            ws.Cells(StockRow, 10).Value = ws.Cells(i, 1).Value
        'Record the TotalVolume
            ws.Cells(StockRow, 11).Value = TotalVolume
        'Increase the StockRow so the next change will be printed on a new line
            StockRow = StockRow + 1
        'Set TotalVolume back to 0 for next Stock
           TotalVolume = 0
        End If
    Next i
Next ws
MsgBox ("Done")
End Sub
