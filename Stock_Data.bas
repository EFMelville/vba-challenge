Attribute VB_Name = "Module1"
Sub Stock_Data()
    
    'Apply code to all worksheets.
    For Each ws In Worksheets
        
        'Establish variables and types.
        Dim TicSym As String
        Dim TotVol As Single
        Dim SumRow As Long
        Dim YrlCh As Double
        Dim PerChg As Double
        Dim FirOp As Double
        
        'Establish variables and types for functionality results.
        Dim GrtVol As Double
        Dim GrtInc As Variant
        Dim GrtDec As Variant
        
        'Establish column headers for results.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Establish column headers for functionality results.
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Establish variable for loop to continue through final row.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        FirOp = ws.Cells(2, 3)
        TotVol = 0
        SumRow = 2
        
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                TicSym = ws.Cells(i, 1).Value
                TotVol = TotVol + ws.Cells(i, 7).Value
                
                ws.Cells(SumRow, 9).Value = TicSym
                ws.Cells(SumRow, 12).Value = TotVol
                
                YrlCh = ws.Cells(i, 6).Value - FirOp
                
                ws.Cells(SumRow, 10).Value = YrlCh
                
                If ws.Cells(SumRow, 10).Value < 0 Then
                    
                    ws.Cells(SumRow, 10).Interior.ColorIndex = 3
                
                Else
                
                    ws.Cells(SumRow, 10).Interior.ColorIndex = 4
                
                End If
                
                '-----------------Space added for personal visual aide-----------------
                
                PerChg = (FirOp - ws.Cells(i, 6).Value) / FirOp
                
                ws.Cells(SumRow, 11).Value = Format(PerChg, "Percent")
                                                
                FirOp = ws.Cells(i + 1, 3)
                SumRow = SumRow + 1
                TotVol = 0
            
            Else
            
                TotVol = TotVol + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        '-----------------Space added for personal visual aide-----------------
        
        GrtVol = ws.Cells(2, 12).Value
        GrtInc = ws.Cells(2, 11).Value
        GrtDec = ws.Cells(2, 11).Value
        
        For i = 2 To LastRow
            
            If ws.Cells(i, 12).Value > GrtVol Then
                
                GrtVol = ws.Cells(i, 12).Value
                
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            Else
                
                GrtVol = GrtVol
            
            End If
            
            '-----------------Space added for personal visual aide-----------------
            
            If ws.Cells(i, 11).Value > GrtInc Then
                
                GrtInc = ws.Cells(i, 11).Value
                
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
            Else
                
                GrtInc = GrtInc
                
            End If
            
            '-----------------Space added for personal visual aide-----------------
            
            If ws.Cells(i, 11).Value < GrtDec Then
                
                GrtDec = ws.Cells(i, 11).Value
                
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
            Else
            
                GrtDec = GrtDec
                
            End If
            
            ws.Cells(2, 17).Value = Format(GrtInc, "Percent")
            ws.Cells(3, 17).Value = Format(GrtDec, "Percent")
            ws.Cells(4, 17).Value = Format(GrtVol, "Scientific")
            
        Next i
        
    Next ws
    
End Sub

