Sub main():
    
    Dim price_o     As Double
    Dim price_c     As Double
    Dim yrChange    As Double
    Dim Ticker      As String
    Dim rowCount    As Integer
    Dim vol         As Double
    Dim lastRow     As Long
    Dim ws          As Worksheet
    
    For Each ws In Worksheets
        rowCount = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        price_o = ws.Cells(2, 3).Value
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                price_c = ws.Cells(i, 6).Value
                Ticker = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                yrChange = price_c - price_o
                
                ws.Range("I" & rowCount).Value = Ticker
                ws.Range("J" & rowCount).Value = yrChange
                Call divZero(ws, rowCount, yrChange, price_o)
                ws.Range("L" & rowCount).Value = vol
                
                Call cellFormats(rowCount, ws)
                
                rowCount = rowCount + 1
                vol = 0
                price_o = ws.Cells(i + 1, 3)
            Else
                
                vol = vol + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        Call greatCal(ws, lastRow)
    Next
End Sub

Sub cellFormats(cnt As Integer, ws As Worksheet):
    If ws.Range("J" & cnt).Value < 0 Then
        
        ws.Range("J" & cnt).Interior.ColorIndex = 3
        
    Else
        
        ws.Range("J" & cnt).Interior.ColorIndex = 4
        
    End If
    ws.Range("I:I", "L:L").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    ws.Range("I:I", "L:L").EntireColumn.AutoFit
    ws.Range("I1:L1").Interior.ColorIndex = 36
End Sub

Sub divZero(ws      As Worksheet, rowCount As Integer, yrChange As Double, open_price As Double)
    
    If open_price = 0 Then
        
        ws.Range("K" & rowCount).Value = 0
        
    Else
        
        ws.Range("K" & rowCount).Value = Format(yrChange / open_price, "Percent")
        
    End If
End Sub

Sub greatCal(ws     As Worksheet, lastRow As Long):
    
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim greatest_vol As Double
    Dim results     As Double
    
    Set Rng = Range("I2:L" & lastRow)
    
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    
    greatest_inc = WorksheetFunction.Max(ws.Range("K:K"))
    greatest_dec = WorksheetFunction.Min(ws.Range("K:K"))
    greatest_vol = WorksheetFunction.Max(ws.Range("L:L"))
    
    ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0))
    ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0))
    ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0))
    
    ws.Cells(2, 17).Value = Format(greatest_inc, "Percent")
    ws.Cells(3, 17).Value = Format(greatest_dec, "Percent")
    ws.Cells(4, 17).Value = greatest_vol
    
    ws.Range("O:O", "Q:Q").EntireColumn.AutoFit
    
End Sub