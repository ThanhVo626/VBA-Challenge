Sub HW2()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim vol As Double
Dim x As Double
Dim j As Double
Dim y As Double
Dim z As Double
Dim k As Double
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
vol = 0
j = 2
y = 0
z = 0
k = 0
ws.Range("I1").Value = "Ticker"
ws.Range("O1").Value = "Ticker"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Value"
For i = 2 To lastrow
    ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
    If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
        If (IsNumeric(ws.Cells(j, 7)) = True And ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
            y = CDbl(ws.Cells(i, 3))
            vol = vol + CDbl(ws.Cells(i, 7).Value)
        ElseIf (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            vol = vol + CDbl(ws.Cells(i, 7).Value)
        End If
    Else
        vol = vol + CDbl(ws.Cells(i, 7).Value)
        ws.Cells(j, 12).Value = vol
        ws.Cells(j, 10).Value = CDbl(ws.Cells(i, 6)) - y
        If ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        ws.Cells(j, 11).Value = 100 * ((CDbl(ws.Cells(i, 6)) / y) - 1)
        If ws.Cells(j, 11).Value > z Then
            z = ws.Cells(j, 11).Value
            ws.Range("O2").Value = ws.Cells(j, 9)
            ws.Range("P2").Value = Str(z) + "%"
        End If
        If ws.Cells(j, 11).Value < k Then
            k = ws.Cells(j, 11).Value
            ws.Range("O3").Value = ws.Cells(j, 9)
            ws.Range("P3").Value = Str(k) + "%"
        End If
        If ws.Cells(j, 12) > ws.Range("P4") Then
            ws.Range("P4").Value = vol
            ws.Range("O4").Value = ws.Cells(j, 9)
        End If
        ws.Cells(j, 11).Value = Str(ws.Cells(j, 11).Value) + "%"
        vol = 0
        j = j + 1
        

    End If
    Next i
Next ws
End Sub