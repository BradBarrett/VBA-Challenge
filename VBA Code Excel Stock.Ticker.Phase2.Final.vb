
Sub StockPhaseTwo()

Dim RNG As Range
Dim RNGVOL As Range
Dim Min As Double
Dim Max As Double
Dim Maxvol As Double
Dim FindRowmin As Long
Dim FindRowmax As Long
Dim Findvolmax As Long
Dim Minticker As String
Dim Maxticker As String
Dim Maxvolticker As String

For Each ws In Worksheets
'Set header
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
'set search range for % chg and vol
Set RNG = ws.Range("M:M")
Set RNGVOL = ws.Range("N:N")
'find min & max & vol
Min = Application.WorksheetFunction.Min(RNG)
Max = Application.WorksheetFunction.Max(RNG)
Maxvol = Application.WorksheetFunction.Max(RNGVOL)
'match the row to min & max to use to return stock ticker
FindRowmax = Application.WorksheetFunction.Match(Max, ws.Range("M:M"), 0)
FindRowmin = Application.WorksheetFunction.Match(Min, ws.Range("M:M"), 0)
Findvolmax = Application.WorksheetFunction.Match(Maxvol, ws.Range("N:N"), 0)
'find stock ticker
Minticker = ws.Range("K" & FindRowmin).Value
Maxticker = ws.Range("K" & FindRowmax).Value
Maxvolticker = ws.Range("K" & Findvolmax).Value
'paste results
ws.Cells(3, 16).Value = Minticker
ws.Cells(2, 16).Value = Maxticker
ws.Cells(4, 16).Value = Maxvolticker
ws.Cells(3, 17).Value = FormatPercent(Min)
ws.Cells(2, 17).Value = FormatPercent(Max)
ws.Cells(4, 17).Value = Maxvol
ws.Cells(4, 17).NumberFormat = "#,##0"


Next ws

End Sub


