Attribute VB_Name = "Module1"
Sub ticker_volume4():

'set some variables

  Dim last_row As Double
  Dim j As Integer
  Dim total As Double
  Dim k As Integer
 
  Dim ws As Worksheet
  
  For Each ws In Worksheets
  
  ws.Range("I1").Value = "Ticker symbol"
  ws.Range("j1").Value = "Total volume"
  ws.Range("k1").Value = "Yearly change"
  ws.Range("l1").Value = "Percent change"
  
  ws.Columns("I:L").AutoFit
  
  last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    
' this loop looks for cells that do match and adds the volumes.  Volumes are in column 7

'Next i

    k = 3
    total = 0

For i = 2 To last_row  'keep adding to total while ticker symbols  match

        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
        total = total + ws.Cells(i, 7).Value
        
        ElseIf ws.Cells(i + 1, 3) <> 0 And ws.Cells(i, 3) <> 0 Then
        
        
        ws.Cells(k, 10).Value = total + ws.Cells(i, 7).Value 'once don't symbols match, place total in column j (adds final volume as well)
         k = k + 1
          total = 0 'reset total
        
        End If
  
Next i

ws.Cells(k, 10) = total + ws.Cells(i, 7).Value 'add last volume to last row

' In this loop, when two values don't match, we calculate he difference in the open (col C) and closing price (col F).

Dim open_price As Double
Dim close_price As Double
Dim change As Double
Dim percent_change As Double

open_price = ws.Cells(2, 3).Value
  
   
   ' row counter for output row
     
    j = 3
  
  '  Loop through all ticker symbols and add new ones to column i

For i = 2 To last_row ' may not pick up last stock because i + 1 after last row would be empty (=0)

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i + 1, 3) <> 0 And ws.Cells(i, 3) <> 0 Then
        
       ws.Cells(j, 9).Value = ws.Cells(i, 1).Value 'add new ticker symbol to next row (column I)
       
       close_price = ws.Cells(i, 6) 'close price is last price in columns 6 before name change
       
       change = close_price - open_price 'will use open price calculated from last round
       
       ws.Cells(j, 11).Value = change ' adds change value to column K.  could add value directly rather than use variable first.  But var makes it more explcit.
       
       percent_change = change / open_price
       
       ws.Cells(j, 12).Value = percent_change
       
       j = j + 1
       open_price = ws.Cells(i + 1, 3) ' sets new open price based on this ticker boundary
       
    End If
    
Next i

    ws.Cells(j, 9).Value = ws.Cells(last_row, 1).Value
    change = ws.Cells(last_row, 6) - open_price
    ws.Cells(j, 11).Value = change
    percent_change = change / open_price
    ws.Cells(j, 12).Value = percent_change

'color percent green (+) or red (-)

last_percent = ws.Cells(Rows.Count, 12).End(xlUp).Row

For i = 2 To last_percent

ws.Cells(i, 12).NumberFormat = "0.00%"

If ws.Cells(i, 11) < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
End If
If ws.Cells(i, 11) > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
End If

Next i

Dim hi_percent As Double
Dim lo_percent As Double
Dim hi_vol As Double

hi_vol = ws.Cells(3, 10)
hi_percent = ws.Cells(3, 12)
lo_percent = ws.Cells(3, 12)

For i = 3 To last_percent

If ws.Cells(i, 10) > hi_vol Then
hi_vol = ws.Cells(i, 10).Value
hi_vol_tick = ws.Cells(i, 9)

End If

If ws.Cells(i, 12) > hi_percent Then
hi_percent = ws.Cells(i, 12).Value
hi_per_tick = ws.Cells(i, 9)
End If

If ws.Cells(i, 12) < lo_percent Then
lo_percent = ws.Cells(i, 12).Value
lo_per_tick = ws.Cells(i, 9)
End If

Next i

'Place output in final summary

ws.Range("N3").Value = "Greatest volume"
ws.Range("N4").Value = "Greatest % increase"
ws.Range("N5").Value = "Greatest % decrease"

ws.Range("O2").Value = "Ticker"
ws.Range("P2").Value = "Value"

ws.Range("O3") = hi_vol_tick
ws.Range("O4") = hi_per_tick
ws.Range("O5") = lo_per_tick

ws.Range("P3") = hi_vol
ws.Range("P4") = hi_percent
ws.Range("P5") = lo_percent

ws.Columns("N:P").AutoFit
ws.Range("P4").NumberFormat = "0.00%"
ws.Range("P5").NumberFormat = "0.00%"

'reset zeros

For i = 2 To last_row
If ws.Cells(i, 3).Value = 1 Then

ws.Cells(i, 3) = 0
ws.Cells(i, 6) = 0

End If

Next i

    
Next ws

End Sub


