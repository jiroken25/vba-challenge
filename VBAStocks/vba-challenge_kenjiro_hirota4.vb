Sub wholeprocess()
Dim WS As Worksheet
' to process everysheet
For Each WS In Worksheets
Application.ScreenUpdating = False
WS.Activate

' to summarise per ticker
Call tickersummary

' to get the greatest value
Call greatest

Next WS
End Sub

Sub tickersummary()
Dim myDic As Object
Dim lastrow As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long
Dim keyticker As Variant
Dim keyopen As Variant
Dim keyclose As Variant
Dim firstdate As Long
Dim lastdate As Long
Dim lastrowticker As Long
' to allow the data more than Long
Dim subtotal As Currency


' count the row of data
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' sort the data by ticker and then by date
ActiveSheet.Sort.SortFields.Add2 Key:=Range(Cells(2, 1), Cells(lastrow, 1)), _
SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(lastrow, 2)), _
SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveSheet.Sort
        .SetRange Range(Cells(1, 1), Cells(lastrow, 9))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

' use dictionary to store the data
Set myDic = CreateObject("Scripting.Dictionary")


' process to each row
' first day of each ticker should be the top because of the soring above.
' So if current row data is different from the data one above, it should be the first data of a ticker.
' last data should be different from the data one below.

i = 2
j = 0
k = 0
For i = 2 To lastrow

subtotal = subtotal + Cells(i, 7).Value


If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
myDic.Add "ticker" & j, Cells(i, 1).Value
myDic.Add "open" & j, Cells(i, 3).Value
j = j + 1
Else
End If
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
myDic.Add "close" & k, Cells(i, 6).Value
myDic.Add "subtotal" & k, subtotal
k = k + 1
subtotal = 0
Else
End If
Next i

' extract data from dictionary for each ticker

For n = 0 To j - 1
keyticker = "ticker" & n
keyopen = "open" & n
keyclose = "close" & n

Cells(n + 2, 9).Value = myDic.Item("ticker" & n)
Cells(n + 2, 10).Value = myDic.Item("close" & n) - myDic.Item("open" & n)

' If open price equal 0, divided cannot work.
If myDic.Item("open" & n) > 0 Then
Cells(n + 2, 11).Value = (myDic.Item("close" & n) / myDic.Item("open" & n)) - 1
Else
Cells(n + 2, 11).Value = "N/A"
End If

' excel function sumifs
Cells(n + 2, 12).Value = myDic.Item("subtotal" & n)
Next n

Set myDic = Nothing

' format settings
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Change%"
Range("L1") = "Total Volume"

Columns("J:J").Select
Selection.Style = "Comma"
Columns("K:K").Select
Selection.Style = "Percent"
Columns("L:L").Select
Selection.NumberFormat = "0"
lastrowticker = Cells(Rows.Count, 9).End(xlUp).Row
Range(Cells(2, 10), Cells(lastrowticker, 10)).Select
Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions(Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions.Count).SetFirstPriority
    With Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions(Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions.Count).SetFirstPriority
    With Range(Cells(2, 10), Cells(lastrowticker, 10)).FormatConditions(1).Interior
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

' header settings
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

' use excel function to get max, min
increase = Application.WorksheetFunction.Max(Range("K:K"))
decrease = Application.WorksheetFunction.Min(Range("K:K"))
total = Application.WorksheetFunction.Max(Range("L:L"))

' identify the ticker number using find
searchkey1 = Format(increase, "0%")
increaserow = Range("K:K").Find(searchkey1, , xlValues).Row
searchkey2 = Format(decrease, "0%")
decreaserow = Range("K:K").Find(searchkey2, , xlValues).Row
searchkey3 = total
totalrow = Range("L:L").Find(searchkey3, , xlValues).Row

' format settings / input the data
Range("O2").Value = Cells(increaserow, 9).Value
Range("P2").Value = increase
Range("P2").Style = "percent"
Range("O3").Value = Cells(decreaserow, 9).Value
Range("P3").Value = decrease
Range("P3").Style = "percent"
Range("O4").Value = Cells(totalrow, 9).Value
Range("P4").Value = total
Range("N:N").ColumnWidth = 20
Range("P:P").ColumnWidth = 30

End Sub





