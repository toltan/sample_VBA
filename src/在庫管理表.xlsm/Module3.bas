Attribute VB_Name = "Module3"
Option Explicit
Public bookName As Workbook

Public Sub WO(ByVal A As String)
Set bookName = Workbooks(A)
End Sub

Public Sub HONDA() 'ホンダリストに入庫を追加したり出庫を引いたり。未完。

Dim ws As Worksheet
Dim i As Range
Dim iRange As Object
Dim VAR As Variant

Set ws = Worksheets("DATA")

    If ws.AutoFilterMode = False Then
    Rows(ws.Range("B3").Row).AutoFilter
    End If
    
ws.Range("B3").AutoFilter FIELD:=2, Criteria1:="2017/4/10" '日付は変数設定

Set iRange = ws.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)

For Each i In iRange

Next
End Sub

Public Sub newPartsNumber()
Dim catNo, oRow, oCol, tRow, tCol, tColZ, i, nRow, bool, inser, iRow, D As Integer
Dim cat, regi As Variant
Dim yesNo3 As String

cat = Array("部品番号", "部品名", "前月残", "前月パレット数")
catNo = 0
tRow = Worksheets("新品番追加").Cells.Find(what:="部品番号").Offset(1, 0).Row
tCol = Worksheets("新品番追加").Cells.Find(what:="部品番号").Offset(1, 0).Column
tColZ = Worksheets("新品番追加").Cells.Find(what:="前月残").Offset(1, 0).Column
bool = 0
inser = 0
D = 0

Worksheets("新品番追加").Activate
Worksheets("新品番追加").Range(Cells(tRow, tCol), Cells(Rows.Count, Columns.Count)).ClearContents



For i = 0 To 3 '在庫表の部品番号、部品名、前月残の行を新品番追加シートに転記
    Worksheets("在庫一覧").Activate
    oRow = Worksheets("在庫一覧").Cells.Find(what:=cat(catNo)).Offset(1, 0).Row
    oCol = Worksheets("在庫一覧").Cells.Find(what:=cat(catNo)).Offset(1, 0).Column
    Worksheets("在庫一覧").Range(Cells(oRow, oCol), Cells(oRow, oCol).End(xlDown)).Copy
    Worksheets("新品番追加").Cells.Find(what:=cat(catNo)).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    catNo = catNo + 1
    
Next

Worksheets("新品番追加").Activate
Worksheets("新品番追加").Range(Cells(tRow, tCol).End(xlDown).Offset(1, 2), Cells(Rows.Count, tColZ).Offset(0, tColZ)).ClearContents
nRow = Worksheets("新品番追加").Cells(Rows.Count, tCol).End(xlUp).Offset(1, 0).Row '新品番追加シートの最終行を取得
  
With Worksheets("新品番追加") '新品番入力

Do While bool >= 0
     
    regi = registration() 'Function registrationより品番、品名を取得

    If regi(0) = "cancel" Then '品番、品名どちらかが未入力の場合はキャンセル
    MsgBox ("キャンセルしました。")
    Exit Sub
    ElseIf regi(0) = "null" Then
    MsgBox ("品番が未入力の為、キャンセルしました。")
    Exit Sub
    End If

    .Cells(nRow + bool, tCol).Value = regi(0) '新品番を追加
    .Cells(nRow + bool, Cells.Find(what:="部品名").Column).Value = regi(1)
    .Cells(nRow + bool, Cells.Find(what:="前月残").Column).Value = "0"
    .Cells(nRow + bool, Cells.Find(what:="前月パレット数").Column).Value = "0"
    
    yesNo3 = MsgBox("続けて新品番登録を行いますか？", vbYesNo + vbQuestion) '２つ以上の新品番があり、連続して登録するならYESでLOOP
    
    If yesNo3 = vbYes Then
    bool = bool + 1
    inser = inser + 1
    Else
    bool = -1
    End If

Loop
    
    If .AutoFilterMode = False Then 'オートフィルターが掛かってなければ掛ける
        .Rows(4).AutoFilter
    End If
    .Range(Cells(tRow - 1, tCol), Cells(Rows.Count, Cells(tRow - 1, tCol).End(xlToRight).Column + 1)).Sort KEY1:=Cells(tRow + 1, tCol), ORDER1:=xlAscending, Header:=xlYes
    '並び替えする

End With

iRow = Worksheets("DATA").Cells.Find(what:="月末在庫").End(xlDown).Offset(1, 0).Row

Worksheets("DATA").Range(iRow & ":" & iRow + inser).Insert 'シート（在庫一覧）、（DATA）に新品番を追加
Worksheets("新品番追加").Range(Cells(tRow, tCol), Cells(Rows.Count, tColZ + 2).End(xlUp)).Copy
Worksheets("DATA").Cells.Find(what:="納品日付").Offset(1, 1).PasteSpecial Paste:=xlPasteValues

    Do Until D = 1 + inser 'シート（DATA）の挿入した行に月末の日にちを挿入した行の分だけ入力
    Worksheets("DATA").Cells.Find(what:="納品日付").End(xlDown).Offset(1, 0).Value = _
    Worksheets("DATA").Cells.Find(what:="納品日付").End(xlDown).Value
    D = D + 1
    Loop

Call stockPlus(D)
Call stockListPlus(D)
Worksheets("新品番追加").Activate
Worksheets("新品番追加").Range(Cells(tRow, tCol), Cells(Rows.Count, tCol).End(xlUp)).Copy
Worksheets("在庫一覧").Cells.Find(what:="部品番号").Offset(1, 0).PasteSpecial Paste:=xlPasteValues
Worksheets("新品番追加").Range(Cells(tRow, Cells.Find(what:="部品名").Column), Cells(Rows.Count, Cells.Find(what:="部品名").Column).End(xlUp)).Copy
Worksheets("在庫一覧").Cells.Find(what:="部品名").Offset(1, 0).PasteSpecial Paste:=xlPasteValues

Call ピボット更新
End Sub

Public Function registration() '新品番登録
Dim A As Boolean
Dim newPart, newPartName, yesNo, yesNo2, oKOnly As String
Dim newP(1) As Variant
A = False

    Do While A = False
    
        newPart = InputBox("新品番を入力してください。")
        
        If newPart = "" Then
            newPart = "null"
            Exit Do
        End If
        
        newPartName = InputBox("新品番の品名を入力してください。")
        
        If newPartName = "" Then
            newPart = "null"
            Exit Do
        End If
        
        yesNo = MsgBox("品番:" & newPart & " " & "品名:" & newPartName & vbCrLf & "以上の内容で登録します。宜しいですか？", vbYesNo)
        
            If yesNo = vbNo Then
            A = False
            yesNo2 = MsgBox("新規品番登録をキャンセルしますか？", vbYesNo)
            
                If yesNo2 = vbYes Then
                newPart = "cancel"
                Exit Do
                End If
                
            Else
            A = True
            End If
            
    Loop
    
newP(0) = newPart
newP(1) = newPartName
registration = newP

End Function

Public Sub stockPlus(ByVal D As Integer) '在庫、納品も行を増やす

Dim lastRow, newLastCol, newLastRow As Integer
Dim i As Integer
Dim newLastRowVal As String
Dim pivotRange As Range
Dim splitPivot As Variant

Worksheets("在庫、納品").Activate

For i = 1 To D

    lastRow = Worksheets("在庫、納品").Cells.Find(what:="項目").End(xlDown).Row 'A列が隠れてる
    newLastCol = Cells(Rows.Count, Cells.Find(what:="部品番号").Column).End(xlUp).Column
    newLastRow = Cells(Rows.Count, Cells.Find(what:="部品番号").Column).End(xlUp).Row
    newLastRowVal = Cells(Rows.Count, Cells.Find(what:="部品番号").Column).End(xlUp).Value
    Set pivotRange = Worksheets("ピボット").Cells.Find(what:="*****")
    
    With Worksheets("在庫、納品")
    
        .Range(lastRow - 5 & ":" & lastRow).Copy
        .Range(lastRow + 1 & ":" & lastRow + 1).PasteSpecial
        
    End With
    
    splitPivot = Split(pivotRange.Address, "$") 'Addressでセル番号を取得し"$"で区切り配列に
    newLastRowVal = "=ピボット!" & splitPivot(1) & newLastRow + 1 '式を入れる
    Worksheets("在庫、納品").Cells(newLastRow + 6, newLastCol).Value = newLastRowVal
    
Next

End Sub

Public Sub stockListPlus(ByVal D As Integer) '在庫一覧も行を増やす Dで回す
Dim noCol, noRow, noEndRow, noEndCol, ist As Integer
Dim i, stockRange, stockRange2 As Range
Dim stockNo, stockNo2, cellNo As Variant
Dim nouba As String
With Worksheets("在庫一覧")
    
    .Activate
    noCol = .Cells.Find(what:="No", LookAt:=xlWhole).Column
    noRow = .Cells.Find(what:="No", LookAt:=xlWhole).Row
    Set stockRange = Worksheets("在庫、納品").Cells.Find(what:="前月残")
    stockNo = Split(stockRange.Address, "$")
    Set stockRange2 = Worksheets("在庫、納品").Cells.Find(what:="合計")
    stockNo2 = Split(stockRange2.Address, "$")
    
End With
    
For ist = 1 To D

    noEndRow = Worksheets("在庫一覧").Cells(noRow, noCol).End(xlDown).Offset(1, 0).Row '在庫一覧の最終行を取得
    noEndCol = Worksheets("在庫一覧").Cells(noRow, Columns.Count).End(xlToLeft).Column '在庫一覧の最終列を取得
    Worksheets("在庫一覧").Rows(noEndRow).Insert

    
    For Each i In Range(Cells(noEndRow, noCol), Cells(noEndRow, noEndCol))
        If i.Offset(-1, 0).End(xlUp).Value = "No" Then
        i.Value = i.Offset(-1, 0).Value + 1
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "前月残" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo(1))
        i.Value = "='在庫、納品'!" & stockNo(1) & cellNo(1) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "前月パレット数" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo(1))
        i.Value = "='在庫、納品'!" & stockNo(1) & cellNo(1) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "入庫累計" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "入庫パレット累計" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "出庫累計" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "出庫パレット累計" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "当日在庫" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "当日パレット数" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='在庫、納品'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "納場" Then
        i.Value = "-"
        End If
        
    Next
    
Next



End Sub





