# VBA-Project
Sub Ticker_Data():
Dim WS As Worksheet
Dim i As Double
Dim j As Double
Dim MyTicker As String
Dim open1 As Double
Dim close1 As Double
Dim volume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim StartRow As Double
Dim EndRow As Double
'Loop through each worksheet
For Each WS In Worksheets
WS.Activate
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
Cells(1, "i").Value = "ticker"
Cells(1, "j").Value = "Yearly Change"
Cells(1, "k").Value = "Percent Change"
Cells(1, "l").Value = "Total Stock Volume"
'Identify the first instance of the Ticker
For i = 2 To lastrow
If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
open1 = Format(Cells(i, 3).Value, "#.00")
StartRow = Cells(i, 1).Row
MyTicker = Cells(i, 1).Value
Cells(Rows.Count, "i").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = MyTicker
End If
'identify the last instance of the Ticker
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
close1 = Format(Cells(i, 6).Value, "#.00")
EndRow = Cells(i, 1).Row
volume = Application.WorksheetFunction.Sum(Range(Cells(StartRow, "g"), Cells(EndRow, "g")))
YearlyChange = close1 - open1
PercentChange = (close1 - open1) / open1
Cells(Rows.Count, "j").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(YearlyChange, "#.00")
Cells(Rows.Count, "k").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(PercentChange, "0.00%")
Cells(Rows.Count, "l").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = volume
End If
Next i
Range("i:l").EntireColumn.AutoFit
lastrow2 = Cells(Rows.Count, "i").End(xlUp).Row
For j = 2 To lastrow2
If Cells(j, "j").Value > 0 Then
Cells(j, "j").Interior.Color = vbGreen
Else
Cells(j, "j").Interior.Color = vbRed
End If
Next j
Range("o2").Value = "Greatest Percent Increase"
Range("o3").Value = "Greatest Percent Decrease"
Range("o4").Value = "Greatest Total Volume"
Range("p1").Value = "ticker"
Range("Q1").Value = "Value"
Range("Q2").Value = Format(Application.WorksheetFunction.max(Range(Cells(2, "k"), Cells(lastrow2, "k"))), "0.00%")
Range("Q3").Value = Format(Application.WorksheetFunction.min(Range(Cells(2, "k"), Cells(lastrow2, "k"))), "0.00%")
Range("Q4").Value = Application.WorksheetFunction.max(Range(Cells(2, "l"), Cells(lastrow2, "l")))
HighIncrease = Format(Range("Q2").Value, "0.00%")
Range("k:k").Find(HighIncrease).Activate
Ticker = ActiveCell.Offset(0, -2).Value
Range("p2").Value = Ticker
HighDecrease = Format(Range("Q3").Value, "0.00%")
Range("k:k").Find(HighDecrease).Activate
Ticker = ActiveCell.Offset(0, -2).Value
Range("p3").Value = Ticker
HighVolume = (Range("Q4").Value)
Range("l:l").Find(HighVolume).Activate
Ticker = ActiveCell.Offset(0, -3).Value
Range("p4").Value = Ticker
Range("O:q").EntireColumn.AutoFit


Next WS

End Sub
