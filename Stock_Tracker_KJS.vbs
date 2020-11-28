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
  
  'reset zeros if necessary
'For i = 2 To last_row
'If ws.Cells(i, 3).Value = 1 Then

'ws.Cells(i, 3) = 0
'ws.Cells(i, 6) = 0

'End If

'Next i

'here I'm changing zeros to 1's to avoid divide by zero error.  stocks should come up as 0 change and 0 % change
For i = 2 To last_row

If ws.Cells(i, 3).Value = 0 Then

ws.Cells(i, 3) = 1
ws.Cells(i, 6) = 1

End If

Next i

  
 ' row counter for output row
  j = 2
 
 '  Loop through all ticker symbols and add new ones to column i
For i = 2 To last_row

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
       ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
       
       j = j + 1
    
    End If
    
' this loop looks for cells that do match and adds the volumes.  Volumes are in column 7

Next i

    k = 2
    total = 0

For i = 2 To last_row  'keep adding to total until ticker symbols don't match

        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
        total = total + ws.Cells(i, 7).Value
       
    Else
         ws.Cells(k, 10).Value = total + ws.Cells(i, 7).Value 'once symbols match, place total in column j (adds final volume as well)
        k = k + 1
        total = 0 'reset total
    End If
        
Next i

' this loop modifies the first loop so that when two values don't match, we calculate he difference in the open (col C) and closing price (col F).

Dim open_price As Double
Dim close_price As Double
Dim change As Double
Dim percent_change As Double

k = 2

open_price = ws.Cells(2, 3).Value

For i = 2 To last_row - 1

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        close_price = ws.Cells(i, 6).Value
        change = close_price - open_price
        percent_change = change / open_price
        ws.Cells(k, 11) = change
        ws.Cells(k, 12) = percent_change
        open_price = ws.Cells(i + 1, 3)
        k = k + 1
        
End If

Next i

' the following is to fill in last row
     
        close_price = ws.Cells(last_row, 6).Value 'get close price directlyfrom last row
        change = close_price - open_price 'open_price should be left over from previous for loop
        percent_change = change / open_price
        ws.Cells(k, 11) = change
        ws.Cells(k, 12) = percent_change

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

hi_vol = ws.Cells(2, 10)
hi_percent = ws.Cells(2, 12)
lo_percent = ws.Cells(2, 12)

For i = 2 To last_percent

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


    



