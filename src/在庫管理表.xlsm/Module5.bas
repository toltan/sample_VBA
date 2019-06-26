Attribute VB_Name = "Module5"
Option Explicit

Public Sub upData() '翌月分の在庫表を作成
Dim oRow, oCol, i, tRow, tCol, tLastRow, tLastCol As Integer
Dim arr, dataArr As Variant
Dim iRange As Range
Dim lastDay, supplier, current, ye, day, yesNo, cate As String

oRow = bookName.Worksheets("在庫一覧").Cells.Find(what:="No", LookAt:=xlWhole).Offset(1, 0).Row
arr = Array("部品番号", "部品名", "当日在庫", "当日パレット数") 'クラウンズポケット
dataArr = Array("部品番号", "部品名", "月末在庫", "月末パレット在庫") 'クラウンズポケット
tRow = bookName.Worksheets("DATA").Cells.Find(what:="納品日付").Offset(1, 0).Row
tCol = bookName.Worksheets("DATA").Cells.Find(what:="納品日付").Column
tLastCol = bookName.Worksheets("DATA").Cells.Find(what:="納品日付").End(xlToRight).Column
ye = Left(Date, 4)
day = Mid(Date, 6, 2) + 1
cate = Worksheets("設定").ComboBox2.Text
If day < 10 Then
day = "0" & day
End If
current = bookName.Worksheets("設定").Range("D3").Text
supplier = bookName.Worksheets("設定").ComboBox1.Text


If wb(current, supplier, ye, day) = 1 Then

yesNo = MsgBox("同名の出荷報告があります。上書きしますか？", vbYesNo)

   If yesNo = vbNo Then
       MsgBox ("出荷報告作成を中断しました。")
       Exit Sub
   End If

End If

lastDay = InputBox("今月末（最終稼働日）の日付を入力してください。" & vbCrLf & "例:2017/5/5  or 5/5")
If lastDay = "False" Then
MsgBox ("出荷報告作成を中断しました。")
Exit Sub
End If
bookName.Worksheets("在庫一覧").Activate

    For i = 0 To 3
    oCol = bookName.Worksheets("在庫一覧").Cells.Find(what:=arr(i), LookAt:=xlWhole).Column
    bookName.Worksheets("在庫一覧").Range(Cells(oRow, oCol), Cells(oRow, oCol).End(xlDown)).Copy
    bookName.Worksheets("DATA").Cells.Find(what:=dataArr(i)).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Next

bookName.Worksheets("DATA").Activate
tLastRow = bookName.Worksheets("DATA").Cells.Find(what:=dataArr(2)).End(xlDown).Row
bookName.Worksheets("DATA").Range(Cells(tLastRow + 1, tCol), Cells(Rows.Count, tLastCol - 3)).ClearContents

    For Each iRange In Range(Cells(tRow, tCol), Cells(tLastRow, tCol)) '日付を入力
    iRange.Value = lastDay
    Next

bookName.Worksheets("在庫一覧").Activate

Application.DisplayAlerts = False
bookName.SaveAs Filename:=current & "\" & cate & "在庫リスト" & ye & "." & day & "月.xlsm"
Application.DisplayAlerts = True

Call ピボット更新
Call stockSheet

End Sub

Public Sub stockSheet()
Dim oRow, oCol, eRow, eCol, tCol, tCol2, tLastCol As Integer
oRow = bookName.Worksheets("ロケーション").Cells.Find(what:="品番").Row
oCol = bookName.Worksheets("ロケーション").Cells.Find(what:="品番").Column
eCol = bookName.Worksheets("ロケーション").Cells.Find(what:="現在庫").Column
tCol = bookName.Worksheets("ロケーション").Cells.Find(what:="前月パレット数").Column
tCol2 = bookName.Worksheets("ロケーション").Cells.Find(what:="出庫日").Column
tLastCol = bookName.Worksheets("ロケーション").Cells.Find(what:="出庫計").Offset(0, -1).Column

With bookName.Worksheets("ロケーション")
.Activate
.Cells(oRow, tCol2).Value = Mid(Date, 6, 2) + 1 & "/01"
eRow = .Cells(Rows.Count, oCol).End(xlUp).Row
.Range(Cells(oRow + 1, eCol), Cells(eRow, eCol)).Copy
.Cells(oRow + 1, tCol).PasteSpecial Paste:=xlPasteValues
.Range(Cells(oRow + 1, tCol2), Cells(Rows.Count, tLastCol)).ClearContents
End With

End Sub

Public Function wb(ByVal current, supplier, ye, day As Variant)
Dim buf, defFile As String
Dim A As Integer
A = 0
buf = Dir(current & "\" & "*.xl*")
defFile = "HONDA" & ye & "." & day & "月.xlsm"
Do While buf <> ""
If buf = defFile Then A = A + 1
buf = Dir()
Loop
wb = A
End Function


