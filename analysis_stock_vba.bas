Attribute VB_Name = "Module1"
Sub stock_data()

'column header
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.UsedRange.Columns.AutoFit

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Next ws


End Sub

Sub tickersymbol()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Range("I2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial

Range("I2").RemoveDuplicates Columns:=1, header:=xlYes

Next ws

End Sub
Sub year()

'yearly change & percent change from opening price at the start of the year
'to closing price at the end of that year. total stock volume.

Dim ws As Worksheet
Dim ticker As String
Dim lastRow As Long
Dim percentyr As Double
Dim openyr As Double
Dim closeyr As Double
Dim changeyr As Double
Dim result As Long
Dim tsvol As Double

tsvol = 0
result = 2

'set worksheet
Set ws = ThisWorkbook.Sheets("2020")

'find last row
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'calculation
For i = 2 To lastRow

If openyr = 0 Then
openyr = ws.Cells(i, 3).Value
End If


If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

closeyr = ws.Cells(i, 6).Value
changeyr = closeyr - openyr
ticker = ws.Cells(i, 1).Value
tsvol = tsvol + ws.Cells(i, 7).Value

If openyr <> 0 Then
percentyr = (changeyr / openyr) * 100
Else
percentyr = 0
End If

'result of formula
ws.Cells(result, 9).Value = ticker
ws.Cells(result, 10).Value = changeyr
ws.Cells(result, 11).Value = percentyr
ws.Cells(result, 12).Value = tsvol

result = result + 1
tsvol = 0
openyr = 0

Else
tsvol = tsvol + ws.Cells(i, 7).Value

End If
Next i

End Sub

Sub percentformat()

Dim ws As Worksheet
Dim lastRow As Long
Dim rng As Range
Dim cell As Range

'set ws
Set ws = ThisWorkbook.Sheets("2020")

'find the lastrow
lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'column to format
Set rng = ws.Range("K2:K" & lastRow)

' Loop through each cell in the range
For Each cell In rng
If Not IsEmpty(cell.Value) Then
cell.Value = cell.Value / 100
cell.NumberFormat = "0.00%"
End If
Next cell

End Sub

Sub color()

Dim ws As Worksheet
Dim J As Range
Dim K As Range
Dim cell As Range
Dim lastRow As Long

'set ws
Set ws = ThisWorkbook.Sheets("2020")

'find the lastrow
lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'column to fill color green & red
Set J = ws.Range("J2:J" & lastRow)
Set K = ws.Range("K2:K" & lastRow)

' Loop through each cell in the range
For Each cell In J
If cell.Value < 0 Then
cell.Interior.color = RGB(255, 0, 0)
Else
cell.Interior.color = RGB(0, 255, 0)
End If
Next cell

For Each cell In K
If cell.Value < 0 Then
cell.Interior.color = RGB(255, 0, 0)
Else
cell.Interior.color = RGB(0, 255, 0)
End If
Next cell

End Sub

Sub greatest()

Dim ws As Worksheet
Dim lastRow As Long
Dim incpercent As Double 'greatest % increase
Dim decpercent As Double 'greatest % decrease
Dim greatvol As Double   'greatest volume
Dim incticker As String 'greatest % increase ticker
Dim decticker As String 'greatest % decrease ticker
Dim volticker As String 'greatest volume ticker

'set worksheet
Set ws = ThisWorkbook.Sheets("2020")
    ws.UsedRange.Columns.AutoFit

'find lastrow
lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

'greatest increase, decrease, volume
incpercent = -1
decpercent = 1
greatvol = 0

'loop through rows
For i = 2 To lastRow

Dim ticker As String
Dim percentch As Double
Dim vol As Double

percentch = ws.Cells(i, 11).Value
ticker = ws.Cells(i, 9).Value
vol = ws.Cells(i, 12).Value

'for greatest increase
If percentch > incpercent Then
incpercent = percentch
incticker = ticker
End If

'for greatest decrease
If percentch < decpercent Then
decpercent = percentch
decticker = ticker
End If

'for greatest volume
If vol > greatvol Then
vol = greatvol
volticker = ticker
End If
Next i

'result table
ws.Cells(2, 17).Value = incpercent
ws.Cells(3, 17).Value = decpercent
ws.Cells(2, 16).Value = incticker
ws.Cells(3, 16).Value = decticker
ws.Cells(4, 16).Value = volticker
ws.Cells(4, 17).Value = vol
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
End Sub


Sub greatpercentformat()

Dim ws As Worksheet
Dim rng As Range
Dim cell As Range

'set ws
Set ws = ThisWorkbook.Sheets("2020")

'column to format
Set rng = ws.Range("Q2:Q4")

' Loop through each cell in the range
For Each cell In rng
If Not IsEmpty(cell.Value) Then
cell.Value = cell.Value / 100
cell.NumberFormat = "0.00%"
End If
Next cell
End Sub
