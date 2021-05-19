Sub vbachallenge()
Dim ws As Worksheet
Set ws = ActiveSheet

'define variables
Dim i, j As Integer

Dim ticker As String
Dim lastrow As Long
Dim vol As Double
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percent As Double


For Each ws In Worksheets


'create table
ws.Cells(1, 9).Value = "Ticker "
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
'lastrow value
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'initial value
j = 2
vol = 0
openprice = ws.Cells(2, 3).Value



For i = 2 To lastrow
'style
ws.Cells(i, 11).Style = "percent"

If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
vol = vol + Cells(i, 7).Value

ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
ws.Cells(j, 12).Value = vol
ws.Cells(j, 9).Value = ticker
closeprice = ws.Cells(i, 6).Value
yearlychange = closeprice - openprice
  percent = yearlychange / (openprice + 0.00000001)
  ws.Cells(j, 10).Value = yearlychange
ws.Cells(j, 11).Value = percent

vol = 0
openprice = ws.Cells(i + 1, 3).Value
j = j + 1

End If


'color index
If ws.Cells(j - 1, 11).Value >= 0 Then
ws.Cells(j - 1, 11).Interior.ColorIndex = 4
 Else: ws.Cells(j - 1, 11).Interior.ColorIndex = 3
ws.Cells(1, 11).Interior.ColorIndex = 0

End If


Next i


'find larger number
ws.Cells(2, 15).Value = "Greatest% Increase"
ws.Cells(3, 15).Value = "Greatest% Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker "

ws.Cells(1, 17).Value = "Value "
Dim gi As Double
Dim gd As Double
Dim gv As Double
Dim x As Integer
Dim tickermax As String
Dim tickermin As String
Dim tickervol As String
Dim lastrow2 As Long

gi = ws.Cells(2, 11).Value
gd = ws.Cells(2, 11).Value
gv = ws.Cells(2, 12).Value
lastrow2 = ws.Range("J65535").End(xlUp).Row

For x = 2 To lastrow2

If ws.Cells(x, 11).Value >= gi Then
gi = ws.Cells(x, 11).Value
tickermax = ws.Cells(x, 9).Value

End If

If ws.Cells(x, 11) <= gd Then
gd = ws.Cells(x, 11).Value
tickermin = ws.Cells(x, 9).Value

End If

If ws.Cells(x, 12) >= gv Then
gv = ws.Cells(x, 12).Value
tickervol = ws.Cells(x, 9).Value

End If

Next x

ws.Cells(2, 16).Value = tickermax

ws.Cells(2, 17).Value = gi
ws.Cells(3, 16).Value = tickermin
ws.Cells(3, 17).Value = gd
ws.Cells(4, 16).Value = tickervol
ws.Cells(4, 17).Value = gv
ws.Cells(2, 17).Style = "percent"
ws.Cells(3, 17).Style = "percent"
Next ws
End Sub
