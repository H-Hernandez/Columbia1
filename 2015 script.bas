Attribute VB_Name = "Module2"
Sub stock_return2()

Dim ticker As String

Dim Vol_total As Double
Vol_total = 0

Dim Summary_row As Integer
Summary_row = 2

Dim Open_price As Double

Dim Close_price As Double

Dim close_price_row As Double
close_price_row = 2

Dim open_price_row As Double
open_price_row = 2

Dim begin_date As Double
begin_date = 20151230

Cells(1, 9).Value = "Ticker summary"
Cells(1, 10).Value = "Volume summary"
Cells(1, 11).Value = "First trading day price"
Cells(1, 12).Value = "Closing trading day price"
Cells(1, 13).Value = "Yearly Change"
Cells(1, 14).Value = "Percent Change"

    Rows("548790:548944").Select
    Selection.Delete Shift:=xlUp

Dim I As Long

For I = 2 To 760037
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value And Cells(I, 3) <> 0 Then
        
        begin_date = 20151230
        
        ticker = Cells(I, 1).Value
        
        Close_price = Cells(I, 6).Value
    
        Vol_total = Vol_total + Cells(I, 7).Value
    
        Range("I" & Summary_row).Value = ticker
    
        Range("J" & Summary_row).Value = Vol_total
        
        Range("L" & close_price_row).Value = Close_price
        
        close_price_row = close_price_row + 1
    
        Summary_row = Summary_row + 1
        
        open_price_row = open_price_row + 1
    
        Vol_total = 0
        
    Else
    
    Vol_total = Vol_total + Cells(I, 7).Value
    
End If
        If Cells(I, 2).Value < begin_date Then
        
            begin_date = Cells(I, 2).Value
            
End If
        If begin_date = Cells(I, 2).Value Then
        
            Open_price = Cells(I, 3).Value
            
            Range("K" & open_price_row).Value = Open_price
            
            
            
End If

Next I

For I = 2 To 3005

    Cells(I, 13).Value = Cells(I, 12).Value - Cells(I, 11).Value
        If Cells(I, 13).Value > 0 Then
            Cells(I, 13).Interior.ColorIndex = 4
        Else
            Cells(I, 13).Interior.ColorIndex = 3
End If

Next I

For I = 2 To 3005

    Cells(I, 14).Value = FormatPercent(((Cells(I, 12).Value / Cells(I, 11).Value) - 1), 2)
    
Next I

End Sub





