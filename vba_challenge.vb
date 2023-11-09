Sub vba_challenge()

    Dim i, j As Integer
    Dim ws As Worksheet
    
    'to loop through each worksheet in the workbook
    For Each ws In Worksheets
    
    
      vol_counter = 0
      ticker_counter = 2
      
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P2").Value = "Ticker"
      ws.Range("Q2").Value = "Value"
      ws.Range("O3").Value = "Greatest % Increase"
      ws.Range("O4").Value = "Greatest % Decrease"
      ws.Range("O5").Value = "Greatest Total Volume"
      
      'defined range to be used in vlookup
      Set myrange = ws.Range("A:C")
        
          
      For i = 2 To ws.Range("A1").End(xlDown).Row
          
          If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
             vol_counter = vol_counter + ws.Cells(i, 7).Value
                  
                  
          Else: ws.Cells(ticker_counter, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(ticker_counter, 10).Value = ws.Cells(i, 6).Value - Application.WorksheetFunction.VLookup(ws.Cells(ticker_counter, 9).Value, myrange, 3, False)
                ws.Cells(ticker_counter, 11).Value = (ws.Cells(ticker_counter, 10).Value / Application.WorksheetFunction.VLookup(ws.Cells(ticker_counter, 9).Value, myrange, 3, False)) * 100
                vol_counter = vol_counter + ws.Cells(i, 7).Value
                ws.Cells(ticker_counter, 12).Value = vol_counter
                vol_counter = 0
                ticker_counter = ticker_counter + 1
          
          End If
          
      Next i
      
      For j = 2 To ws.Range("J1").End(xlDown).Row
          
          If ws.Cells(j, 10).Value > 0 Then
             ws.Cells(j, 10).Interior.ColorIndex = 4
             
          Else:
              ws.Cells(j, 10).Interior.ColorIndex = 3
          
          End If
          
      Next j
      
      ws.Range("Q3").Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
      ws.Range("Q4").Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
      ws.Range("Q5").Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
      
      'use the combination of index and match to get the 'Ticker' based on the 'Value'
      ws.Range("P3").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
      ws.Range("P4").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("K:K"), 0))
      ws.Range("P5").Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(ws.Range("Q5").Value, ws.Range("L:L"), 0))
    
    Next ws
    

End Sub
