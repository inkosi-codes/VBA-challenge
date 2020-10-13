Sub output():
    
    Dim price_o As Double
    Dim price_c As Double
    Dim yrChange As Double
    Dim ticker As String
    Dim rowCount As Integer
    Dim vol As Double
    Dim lastRow As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        rowCount = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        price_o = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                price_c = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                yrChange = price_c - price_o
                
                ws.Range("I" & rowCount).Value = ticker
                ws.Range("J" & rowCount).Value = yrChange
                Call divZero(ws, rowCount, yrChange, price_o)
                ws.Range("L" & rowCount).Value = vol
                
                Call colorChange(rowCount, ws)
                
                rowCount = rowCount + 1
                vol = 0
                price_o = ws.Cells(i + 1, 3)
            Else
                
                vol = vol + ws.Cells(i, 7).Value
                
            End If
            Next i
        Next
        
    End Sub
    
    Sub colorChange(cnt As Integer, ws As Worksheet):
        If ws.Range("J" & cnt).Value < 0 Then
            
            ws.Range("J" & cnt).Interior.ColorIndex = 3
            
        Else
            
            ws.Range("J" & cnt).Interior.ColorIndex = 4
            
        End If
        
        ws.Range("I" & cnt).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ws.Range("J" & cnt).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ws.Range("K" & cnt).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ws.Range("L" & cnt).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End Sub
    
    Sub divZero(ws As Worksheet, rowCount As Integer, yrChange As Double, open_price As Double)
        
        If open_price = 0 Then
            
            ws.Range("K" & rowCount).Value = 0
            
        Else
            
            ws.Range("K" & rowCount).Value = yrChange / open_price
            
        End If
    End Sub
    

