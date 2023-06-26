Sub Stockyear()
' Declare Current as a worksheet object variable.
Dim ws As Worksheet
 ' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets


'Define variables
Dim change_in_Price As Double
Dim i As Long
Dim j As Long
Dim Lastrow As Long
Dim Start As Long
Dim Greatestincrease As Double
Dim Greatestdecrease As Double
Dim GreatestTotalVolume As LongLong


'determine the last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'determine the summary last row
SummaryLastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Set intial values
j = 0
Percent_change = 0
Change = 0
Total_Stock_Volume = 0
Start = 2
ws.Range("I1").Value = "Ticker_Name"
ws.Range("J1").Value = "Change"
ws.Range("K1").Value = "Percent_Change"
ws.Range("L1").Value = "Total Stock Volume"



ws.Range("O2").Value = "Greatest % Increase"

ws.Range("O3").Value = "Greatest % Decrease"

ws.Range("O4").Value = "Greatest Total Volume"


ws.Range("P1").Value = "Ticker"

ws.Range("Q1").Value = "Value"

'Write loop
For i = 2 To Lastrow
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    Change = ws.Cells(i, 6) - ws.Cells(Start, 3)
    Percent_change = Change / ws.Cells(Start, 3)
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
    Start = i + 1
    
    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
    ws.Range("J" & 2 + j).Value = Change
    'Multi-conditional formatting
        If ws.Range("J" & 2 + j).Value > 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & 2 + j).Value < 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        End If

    ws.Range("K" & 2 + j).Value = Percent_change
    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
    ws.Range("L" & 2 + j).Value = Total_Stock_Volume
    Total_Stock_Volume = 0
    j = j + 1
    
    Else
    
     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)

    
End If


Next i

'Finding the greatest increase and decrease

Greatestincrease = WorksheetFunction.Max(ws.Range("K2:K" & SummaryLastrow))
ws.Cells(2, 17).Value = Greatestincrease
ws.Range("Q2").NumberFormat = "0.00%"
Greatestdecrease = WorksheetFunction.Min(ws.Range("K2:K" & SummaryLastrow))
ws.Cells(3, 17).Value = Greatestdecrease
ws.Range("Q3").NumberFormat = "0.00%"
GreatestTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & SummaryLastrow))
ws.Cells(4, 17).Value = GreatestTotalVolume

'Using match to lookup the value
Greatestincreaserow = WorksheetFunction.Match(Greatestincrease, ws.Range("K2:K" & SummaryLastrow), 0)
ws.Range("P2") = ws.Cells(Greatestincreaserow + 1, 9).Value

Greatestdecreaserow = WorksheetFunction.Match(Greatestdecrease, ws.Range("K2:K" & SummaryLastrow), 0)
ws.Range("P3") = ws.Cells(Greatestdecreaserow + 1, 9).Value

Greatesttotalvolumerow = WorksheetFunction.Match(GreatestTotalVolume, ws.Range("L2:L" & SummaryLastrow), 0)
ws.Range("P4") = ws.Cells(Greatesttotalvolumerow + 1, 9).Value

Next ws

End Sub