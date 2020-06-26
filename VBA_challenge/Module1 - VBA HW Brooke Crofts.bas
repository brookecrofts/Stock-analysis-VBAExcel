Attribute VB_Name = "Module1"
Sub Final_HW()

'Inserting Titles Via Ranges
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Stock_Volume"

'Setting variables
    Dim ticker As String
    Dim stock_volume As Double
        stock_volume = 0
    Dim table As Integer
        table = 2
    Dim open_price As Double
        open_price = Range("C2")
    Dim close_price As Double
    Dim change As Double
    Dim percent As Double
    
 'Setting up Loop for all four columns
    For i = 2 To Range("A1").CurrentRegion.End(xlDown).Row
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 3).Value <> 0 Then
        ticker = Cells(i, 1).Value
        stock_volume = stock_volume + Cells(i, 7)
        Range("I" & table).Value = ticker
        Range("L" & table).Value = stock_volume
        table = table + 1
        stock_volume = 0
        open_price = Cells(i + 1, 3).Value
        
        Else
        stock_volume = stock_volume + Cells(i, 7).Value
        close_price = Cells(i + 1, 6).Value
        change = close_price - open_price
        Range("J" & table).Value = change
        percent = (close_price / open_price) - 1
        Range("K" & table).Value = percent
        Range("K" & table).NumberFormat = "0.00%"
        End If

    'Color the positive green and negative red
        If Range("J" & table).Value >= 0 Then
           Range("J" & table).Interior.ColorIndex = 4
        Else
           Range("J" & table).Interior.ColorIndex = 3
        End If
     Next i
End Sub
