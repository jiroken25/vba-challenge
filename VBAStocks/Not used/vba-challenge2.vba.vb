Sub wholeprocess()
Dim WS As Worksheet
Application.ScreenUpdating = False

Call ticker
Call yearchange
Call greatest

End Sub


Sub ticker()
Dim myDic As Object
Dim buf As Variant
Dim WS As Worksheet
Dim lastrow As Long

Dim i As Long
Dim j As Long

Dim count As Long


Set myDic = CreateObject("Scripting.Dictionary")
lastrow = Cells(Rows.count, 1).End(xlUp).Row

i = 1
For i = 1 To lastrow

buf = Cells(i, 1).Value
If Not myDic.Exists(buf) Then
       myDic.Add buf, buf
End If
Next i

items = myDic.items

j = 0
For j = 0 To myDic.count - 1
Cells(j + 1, 9) = items(j)

Next j

Set myDic = Nothing




End Sub


Sub yearchange()
Dim lastrowdata As Long
Dim lastrowticker As Long
Dim i As Long
Dim j As Long
Dim firstdate As Long
Dim lastdate As Long
Dim searchkey1 As Variant
Dim searchkey2 As Variant
Dim firstvalue As Double
Dim lastvalue As Double
Dim findkey As Range
Dim first As String
Dim last As String



Columns("J:J").Select
Selection.Style = "Comma"
Columns("K:K").Select
Selection.Style = "Percent"
Columns("L:L").Select
Selection.NumberFormat = "0"
Selection.ColumnWidth = 30

lastrowdata = Cells(Rows.count, 1).End(xlUp).Row
lastrowticker = Cells(Rows.count, 9).End(xlUp).Row

fisrtdate = Application.WorksheetFunction.Min(Range("B:B"))
lastdate = Application.WorksheetFunction.Max(Range("B:B"))
searchkey = firstdate

For i = 2 To lastrowticker

searchkey1 = firstdate
Set findkey = Range("B:B").Find(searchkey1, , xlValues)
first = findkey.Address
If Cells(findkey.Row, 1).Value = Cells(i, 9).Value Then
firstvalue = Cells(findkey.Row, 3)
Else

Do
Set findkey = Range("B:B").FindNext(after:=findkey)
If findkey = first Then
Exit Do
Else
If Cells(findkey.Row, 1) = Cells(i, 9) Then
firstvalue = Cells(findkey.Row, 3)
Exit Do
Else
End If
End If

Loop
End If

searchkey2 = lastdate
Set findkey = Range("B:B").Find(searchkey2, , xlValues)
last = findkey.Address
If Cells(findkey.Row, 1).Value = Cells(i, 9).Value Then
lastvalue = Cells(findkey.Row, 6)

Else

Do
Set findkey = Range("B:B").FindNext(after:=findkey)
If findkey.Address = last Then
Exit Do
Else
If Cells(findkey.Row, 1) = Cells(i, 9) Then
lastvalue = Cells(findkey.Row, 6)
Exit Do
Else
End If
End If

Loop
End If



Cells(i, 10).Value = lastvalue - firstvalue
Cells(i, 11).Value = lastvalue / firstvalue - 1
Cells(i, 12).Value = Application.WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), Cells(i, 9).Value)

Next i

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Change%"
Range("L1") = "Total Volume"

Range(Cells(2, 10), Cells(lastrowticker, 10)).Select
Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With



End Sub

Sub greatest():
Dim increase As Double
Dim decrease As Double
Dim total As Double
Dim increaserow As Long
Dim decreaserow As Long
Dim totalrow As Double
Dim searchkey1 As Variant
Dim searchkey2 As Variant
Dim searchkey3 As Variant


Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

increase = Application.WorksheetFunction.Max(Range("K:K"))
decrease = Application.WorksheetFunction.Min(Range("K:K"))
total = Application.WorksheetFunction.Max(Range("L:L"))

searchkey1 = Format(increase, "0%")
increaserow = Range("K:K").Find(searchkey1, , xlValues).Row

searchkey2 = Format(decrease, "0%")
decreaserow = Range("K:K").Find(searchkey2, , xlValues).Row

searchkey3 = total


totalrow = Range("L:L").Find(searchkey3, , xlValues).Row

Range("O2").Value = Cells(increaserow, 9).Value
Range("P2").Value = increase
Range("P2").Style = "percent"
Range("O3").Value = Cells(decreaserow, 9).Value
Range("P3").Value = decrease
Range("P3").Style = "percent"
Range("O4").Value = Cells(totalrow, 9).Value
Range("P4").Value = total
Range("P:P").ColumnWidth = 30

End Sub

