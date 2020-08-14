Attribute VB_Name = "Module1"
Sub stonks():

'ws Variables
    
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "% Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P1").Value = "Ticker Symbol"
        ws.Range("Q1").Value = "Value"

'Stocks Variables
        
        Dim i As Long
        Dim j As Integer
        Dim labelrow As Long
        Dim sumrow As Long
        Dim ticker As String
        Dim opening As Double
        Dim closing As Double

'Stock equation setups
        
        labelrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        sumrow = 2
        opening = ws.Cells(2, 3).Value

        For i = 2 To labelrow
            
            'Total Stock Volume
                ws.Cells(sumrow, 12).Value = ws.Cells(sumrow, 12).Value + ws.Cells(i, 7).Value
                
            'Ticker Symbol
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Cells(sumrow, 9).Value = ticker
                         
            'Yearly Change
                closing = ws.Cells(i, 6).Value
                ws.Cells(sumrow, 10).Value = closing - opening
                
            'Percent Change
                If opening <> 0 Then
                    ws.Cells(sumrow, 11).Value = (closing - opening) / opening
                Else
                    ws.Cells(sumrow, 11).Value = 0
                End If
                
            'Color coding
                If ws.Cells(sumrow, 10).Value >= 0 Then
                    ws.Cells(sumrow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(sumrow, 10).Interior.ColorIndex = 3
                End If
                
            'Move onto next rows
                sumrow = sumrow + 1
                opening = ws.Cells(i + 1, 3).Value
                End If
       
'Outputs and Formatting
    Next i
        ws.Range("K2:K" & sumrow).NumberFormat = "0.00%"
        ws.Range("L2:L" & sumrow).NumberFormat = "#,##0"
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & sumrow))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & sumrow))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & sumrow))
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,##0"
        ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & sumrow), 0))
        ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & sumrow), 0))
        ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I2:I" & sumrow), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & sumrow), 0))
        ws.Columns.AutoFit
    
'Repeat for remaining worksheets
    Next ws

End Sub

