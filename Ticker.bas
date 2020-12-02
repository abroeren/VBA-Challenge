Attribute VB_Name = "Module1"
Sub Ticker()

Dim ws As Worksheet
    
For Each ws In ActiveWorkbook.Worksheets

        'declarations
        Dim Ticker As String
        
        Dim SumRow As Integer
            SumRow = 2
        
        Dim Tickerstart As Double
            Tickerstart = 2
        
        Dim StockOpen As Double
            StockOpen = 0
        
        Dim StockClose As Double
            StockClose = 0
        
        Dim Volume As Double
            Volume = 0
        
        LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'columnheadings
        ws.Range("J1").FormulaR1C1 = "Ticker"
        ws.Range("K1").FormulaR1C1 = "Yearly Change"
        ws.Range("L1").FormulaR1C1 = "Percent Change"
        ws.Range("M1").FormulaR1C1 = "Total Stock Volume"
            
        'loop through worksheet
            For I = 2 To LRow
                
                If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
           
                'find ticker
                Ticker = ws.Cells(I, 1).Value
                'find first row in ticker stock open
                StockOpen = ws.Cells(Tickerstart, 3).Value
                'find last row in ticker stock close
                StockClose = ws.Cells(I, 6).Value
                'deduct stock close open from stock close
                yearlyChange = StockClose - StockOpen
                
                    'calculate percentage change between stock open and stock close
                    If StockOpen > 0 Then
                    percentChange = (StockClose - StockOpen) / StockOpen
                    Else
                    percentChange = "N/A"
                    End If
                                        
                'calculate total volume
                Volume = Volume + ws.Cells(I, "g").Value
                
                'populate the summary table
                ws.Range("J" & SumRow).Value = Ticker
                ws.Range("K" & SumRow).Value = yearlyChange
                ws.Range("L" & SumRow).Value = percentChange
                    
                    'Color coding
                    If yearlyChange < 0 Then
                        ws.Range("K" & SumRow, "L" & SumRow).Interior.Color = RGB(255, 0, 0)
                        Else
                        ws.Range("K" & SumRow, "L" & SumRow).Interior.Color = RGB(0, 255, 0)
                        End If
                            
                    'Percentage formatting
                    ws.Range("L" & SumRow).Value = percentChange
                    ws.Range("L" & SumRow).NumberFormat = "0.00%"
                                
                ws.Range("m" & SumRow).Value = Volume
                
                'value resets
                SumRow = SumRow + 1
                StockOpen = 0
                StockClose = 0
                Volume = 0
                Tickerstart = I + 1
                
            'next ticker
            
            Else
                Volume = Volume + ws.Cells(I, "g").Value
            
            End If
        
        Next I
    
    'next worksheet
    Next ws
    
End Sub
