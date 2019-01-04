# visual-basic

               
Sub Stock_Volume()
    
    Dim Ticker_Symbol As String
    Dim Volume_Total As Double
    Dim Last_Row As Double
    Dim Volume_Table As Integer
    
      For Each ws In Worksheets
          ws.Range("K1").Value = "Ticker_Symbol"
          ws.Range("L1").Value = "Volume_Traded"
          Volume_Total = 0
          Volume_Table = 2
          Last_Row = ws.Range("A1").End(xlDown).Row
          
          For i = 2 To Last_Row
    
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
              Ticker_Symbol = ws.Cells(i, 1).Value
    
              Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
              ws.Range("K" & Volume_Table).Value = Ticker_Symbol
    
              ws.Range("L" & Volume_Table).Value = Volume_Total
    
              Volume_Table = Volume_Table + 1
              
              Volume_Total = 0
    
              Else
    
              Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
              End If
    
          Next i
      Next ws
End Sub




