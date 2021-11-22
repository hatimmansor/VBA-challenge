Attribute VB_Name = "Module1"
Sub Task1()
' Create a variable to hold ticker symbol, year open and close price and total stock

 Dim ticker_symbol As String
 Dim open_price, close_price, total_stock As Double
 Dim result_counter As Integer
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
For Each ws In Worksheets
   
    ' activate the selected work sheet
    Worksheets(ws.Name).Activate
    
    
     'initilize some varuibles
     ticker_symbol = Cells(2, 1).Value
     open_price = Cells(2, 3).Value
     total_stock = Cells(2, 7).Value
     result_counter = 0
     
     ' print the new table header
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Yearly Change"
     Cells(1, 11).Value = "Percent Change"
     Cells(1, 12).Value = "Total Stock Volume"
     
     ' apply the conditional formating
      Columns("K:K").Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
     
     
     'Loop through all the stocks from numbers 2 till the end of column
     
    For i = 2 To Range("A1").End(xlDown).Row
        ' check if other ticker has been detected
        If (Cells(i, 1).Value = ticker_symbol) Then
        
            total_stock = total_stock + Cells(i, 7).Value
        Else
        ' when othe ticker found, then do the colculation and print the results on the same sheet
            close_price = Cells(i - 1, 6).Value
            result_counter = result_counter + 1
            Cells(result_counter + 1, 9).Value = ticker_symbol
            Cells(result_counter + 1, 10).Value = close_price - open_price
            If open_price <> 0 Then
                Cells(result_counter + 1, 11).Value = (close_price - open_price) / open_price
            Else
                Cells(result_counter + 1, 11).Value = 0
            End If
            
            Cells(result_counter + 1, 12).Value = total_stock
            
            ticker_symbol = Cells(i, 1).Value
            open_price = Cells(i, 3).Value
            total_stock = Cells(i, 7).Value
        End If
          
     Next i
     
     
      ' print the bonus part table header
      ' print the new table header
     Cells(1, 16).Value = "Ticker"
     Cells(1, 17).Value = "Value"
     Cells(2, 15).Value = "Greatest % Increase"
     Cells(3, 15).Value = "Greatest % Decrease"
     Cells(4, 15).Value = "Greatest Total Volume"
     
     ' calculate the bonus part Variable deceleration
     Dim find_value As Double
     Dim find_ticker As String
     ' find the gretest %increase
     
     find_value = Cells(2, 11).Value
     find_ticker = Cells(2, 9).Value
     
     For i = 2 To Range("k1").End(xlDown).Row
        If Cells(i, 11).Value > find_value Then
            find_value = Cells(i, 11).Value
            find_ticker = Cells(i, 9).Value
        End If
     Next i
     Cells(2, 16).Value = find_ticker
     Cells(2, 17).Value = find_value
     ' find the gretest %Decrease
     
     find_value = Cells(2, 11).Value
     find_ticker = Cells(2, 9).Value
     
     For i = 2 To Range("k1").End(xlDown).Row
        If Cells(i, 11).Value < find_value Then
            find_value = Cells(i, 11).Value
            find_ticker = Cells(i, 9).Value
        End If
     Next i
     Cells(3, 16).Value = find_ticker
     Cells(3, 17).Value = find_value
    
     find_value = Cells(2, 12).Value
     find_ticker = Cells(2, 9).Value
     
     For i = 2 To Range("k1").End(xlDown).Row
        If Cells(i, 12).Value > find_value Then
            find_value = Cells(i, 12).Value
            find_ticker = Cells(i, 9).Value
        End If
     Next i
     Cells(4, 16).Value = find_ticker
     Cells(4, 17).Value = find_value
Next ws

End Sub
