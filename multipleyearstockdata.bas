Attribute VB_Name = "Module1"
Sub multipleyearstockdata()
    
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Dim i As Long
        Dim b As Long
        Dim ticker As String
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim volume As Double
        volume = 0
        Dim r As Double
        r = 2
        Dim opening As Double
        Dim closing As Double

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
       opening = ws.Cells(2, 3).Value
       
        For i = 2 To lastrow
         
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              
                ticker = ws.Cells(i, 1).Value
                ws.Cells(r, 9).Value = ticker
               
                closing = ws.Cells(i, 6).Value
               
                yearlychange = closing - opening
                ws.Cells(r, 10).Value = yearlychange
                    
                    If ws.Cells(r, 10).Value > 0 Then
                        ws.Cells(r, 10).Interior.Color = RGB(0, 255, 0)
                    
                    Else
                        ws.Cells(r, 10).Interior.Color = RGB(255, 0, 0)
                    End If
    
                    If (opening = 0 And closing = 0) Then
                        percentchange = 0
                    
                    ElseIf (opening = 0 And closing <> 0) Then
                        percentchange = 1
                    
                    Else
                        percentchange = yearlychange / opening
                        ws.Cells(r, 11).Value = percentchange
                        ws.Cells(r, 11).NumberFormat = "0.00%"
                    End If
                
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(r, 12).Value = volume
               
                r = r + 1
                
                opening = ws.Cells(i + 1, 3)
               
                volume = 0
            
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        lastrowB = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For b = 2 To lastrowB
            If ws.Cells(b, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowB)) Then
                ws.Cells(2, 16).Value = Cells(b, 9).Value
                ws.Cells(2, 17).Value = Cells(b, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(b, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowB)) Then
                ws.Cells(3, 16).Value = Cells(b, 9).Value
                ws.Cells(3, 17).Value = Cells(b, 11).Value
               ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(b, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowB)) Then
                ws.Cells(4, 16).Value = Cells(b, 9).Value
                ws.Cells(4, 17).Value = Cells(b, 12).Value
            End If
        
        Next b
        
        ws.Columns.AutoFit
        
    Next ws
        
End Sub

