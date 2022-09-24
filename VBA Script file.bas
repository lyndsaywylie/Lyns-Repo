Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
   'Setting Header Names to summarized data
    Dim Ticker As String
    Range("I1") = "Ticker"
    Dim Headers As Variant
    Headers = Array("Yearly Change", "Percent Change", "Total Stock Volume")
    Range("J1:L1") = Headers
    
    'Keep track of location for each ticker in table
    Dim Ticker_table As Long
    Ticker_table = 2
    
    'Set variables for holding i, LastRow, TickerName, Yearly Change, Percent Change, Total Stock Volume, End Price & Start Price
    Dim i As Long, LastRow As Long, TickerName As String, Yearly_Change As Double, Percent_Change As Long, TotalVol As LongPtr, Start_Price As Variant, End_Price As Double

    'Set variable for Open Price, Close Price, Price Change & Price Percent
    Dim Prices As Variant
    Prices = Array("Open_Price", "Close_Price", "Price_Change", "Percent_Change")
    
    'Determine the lastrow
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set all appropriate variables to 0
    TotalVol = 0
    Open_Price = 0
    Close_Price = 0
    Price_Change = 0
    Price_Percent = 0

    'Set Price for first ticker's open value
    
Open_Price = Cells(2, 3).Value
  
    For i = 2 To LastRow
        
         'Check to see if still within same ticker, if not..
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the Ticker name & Close Price
            TickerName = Cells(i, 1).Value
            Close_Price = Cells(i, 6).Value
        
            'Calculate Price Change, Percent Change & Volume
            Price_Change = Close_Price - Open_Price
            Percent_Change = (Price_Change / Open_Price) * 100
            TotalVol = TotalVol + Cells(i, 7).Value
               
            'Print ticker name, price change, percent change & volume change in table
            Range("i" & Ticker_table).Value = TickerName
            Range("j" & Ticker_table).Value = Price_Change
            Range("k" & Ticker_table).Value = Percent_Change
            Range("l" & Ticker_table).Value = TotalVol
        
        'Assign colors to Yearly change
        If (Price_Change > 0) Then
            Range("j" & Ticker_table).Interior.ColorIndex = 4
        ElseIf (Price_Change < 0) Then
            Range("j" & Ticker_table).Interior.ColorIndex = 3
        End If
                      
        'Adjust column width
        Range("I:I").ColumnWidth = 8.43
        Range("J:J").ColumnWidth = 13.14
        Range("K:K").ColumnWidth = 14.29
        Range("L:L").ColumnWidth = 17.57
                
        'Update Percent Change format
        Range("K" & Ticker_table).Value = (CStr(Percent_Change) & "%")
               
        'Reset Price Change & Percent Change
        Price_Change = 0
        Percent_Change = 0
        
         'Add new row
        Ticker_table = Ticker_table + 1
         
        'Reset Close Price, Open Price & Price Change
        Close_Price = 0
        Open_Price = Cells(i + 1, 3).Value
        Price_Change = 0
                
        'Reset the volume total
        TotalVol = 0
        
        'If the cell immmediately following row is the same
        Else
        
        TotalVol = TotalVol + Cells(i, 7).Value
    
        
    End If
Next i

End Sub
