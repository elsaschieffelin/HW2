Attribute VB_Name = "Module1"
Sub VBHomeworkModerate():

For Each ws In Worksheets

'Name and fit additional column
    ws.Range("J1").Value = "Ticker"
        ws.Columns("J:J").AutoFit
    ws.Range("K1").Value = "Yearly Change"
        ws.Columns("K:K").AutoFit
    ws.Range("L1").Value = "Precent Change"
        ws.Columns("L:L").AutoFit
    ws.Range("M1").Value = "Total Volume Change"
        ws.Columns("M:M").AutoFit


'Set Variables
    Dim StockRow As Long
    Dim TotalVolume As Double
    Dim PriceChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    

'Set StockRow equal to 2 so we can start printing tickers in a clean list and TotalVolume, OpenPrice and ClosePrice equal to 0
    StockRow = 2
    TotalVolume = 0
    OpenPrice = 0
    ClosePrice = 0
'Find the Last Row
    Dim LastRow As Long

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'identify the change in stock and add up the total stock volume for each stock in 2016
    Dim i As Long
    For i = 2 To LastRow
        
        'Begin to tally up the total stock volume
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        'Add up OpenPrice and ClosePrice
        OpenPrice = OpenPrice + ws.Cells(i, 3).Value
        ClosePrice = ClosePrice + ws.Cells(i, 6).Value
        
        'Identify where the stock ticker changes
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Record ticker
            ws.Cells(StockRow, 10).Value = ws.Cells(i, 1).Value
            'Record the TotalVolume
            ws.Cells(StockRow, 13).Value = TotalVolume
            'calculate and record price change
            ws.Cells(StockRow, 11).Value = ClosePrice - OpenPrice
            
            'Calculate the percent change in price
            ws.Cells(StockRow, 12).Value = FormatPercent((ClosePrice - OpenPrice) / OpenPrice)
        
            'Increase the StockRow so the next change will be printed on a new line
            StockRow = StockRow + 1
        'Set TotalVolume and OpenPrice and ClosePrice back to 0 for next Stock
           TotalVolume = 0
           OpenPrice = 0
           ClosePrice = 0
        End If
    Next i

'formatting
    Dim ChangeLastRow As Long
    ChangeLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

    Dim j As Integer
    For j = 2 To ChangeLastRow
              'format coloring
            If ws.Cells(j, 11).Value >= 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
          End If
    Next j
Next ws
MsgBox ("Done")
End Sub
