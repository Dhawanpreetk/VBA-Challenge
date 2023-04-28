
Sub Stockticker():

 Dim ws As Worksheet
 
 
 For Each ws In Worksheets
 Dim Ticker_Name As String
 Dim yearly_change As Double
 Dim Percentage_change As Double
 Dim totalstock_volume As Single
 Dim opening_price  As Double
 Dim closing_price As Double
 Dim startprice As Long
 
 
  Dim increase_ticker As String
  
  Dim decrease_ticker As String
  
  Dim greatestincrease As Double
  
  Dim volume As Double
  
  
  
     'Dim LastRow As Long
  startprice = 2
  totalstock_volume = 0
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  'Determine the last row'
  LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
 'Adding titles to the columns'
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percentage Change"
 ws.Range("L1").Value = "Total Stock Volume"
  
  
 'loops that goes into rows and columns to pick ticker name'
For i = 2 To LastRow
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       totalstock_volume = totalstock_volume + ws.Cells(i, 7).Value

        yearly_change = ws.Cells(i, 6).Value - ws.Cells(startprice, 3).Value
        
        
        
        
        Percentage_change = yearly_change / ws.Cells(startprice, 3).Value
        
        
        
        'Printing the values in respective columns'
        
        ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1).Value
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
        ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
        ws.Range("K" & Summary_Table_Row).Value = Percentage_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        ws.Range("L" & Summary_Table_Row).Value = totalstock_volume
         
         
         If yearly_change <= 0 Then
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
         Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
        
        
        
      'Then goes to the next row and prints another ticker name'
         Summary_Table_Row = Summary_Table_Row + 1
        ' Reset the
          totalstock_volume = 0
          yearly_change = 0
          Percentage_change = 0
          startprice = i + 1
    Else
    
    
     totalstock_volume = totalstock_volume + ws.Cells(i, 7).Value
    End If
    
     Next i
    'Printing the Column names '
    
    ws.Cells(3, 15).Value = "Greatest%Increase"
    ws.Cells(4, 15).Value = "Greatest%Decrease"
    ws.Cells(5, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 16).Value = "Ticker"
    ws.Cells(2, 17).Value = "Value"
    
    'Calculating the Values '
    
    ws.Cells(3, 17).Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
    ws.Cells(4, 17).Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
    ws.Cells(5, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) * 100
    
    
     
    
    
        
    
    
    
   
   Next ws
End Sub

