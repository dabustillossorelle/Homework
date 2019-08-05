Sub multi_yr_esy()
Dim ws As Worksheet
Dim Ticker_Name As String
 
Dim Stock_Volume_Tot As Double
    Stock_Volume_Tot = 0
 
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
For Each ws In Worksheets
Summary_Table_Row = 2

 For i = 2 To 760192
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
     Ticker_Name = ws.Cells(i, 1).Value
     
     Stock_Volume_Tot = Stock_Volume_Tot + ws.Cells(i, 7).Value
    
     ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
   
     ws.Range("K" & Summary_Table_Row).Value = Stock_Volume_Tot
    
     Summary_Table_Row = Summary_Table_Row + 1

     Stock_Volume_Tot = 0
  
   Else
    
     Stock_Volume_Tot = Stock_Volume_Tot + ws.Cells(i, 7).Value
   End If
 Next i
 
 ws.Cells(1, 10).Value = "Ticker Name"
 ws.Cells(1, 11).Value = "Stock Volume Total"
 
 Next ws
 
End Sub


