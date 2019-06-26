Attribute VB_Name = "Module1"
Option Explicit

'会社では64ビット対応にする Ptrsafe & longptr
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpCaptionName As String) As LongPtr
Declare PtrSafe Sub SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr)

Public exImCount(3) As Integer

Public Sub callUserForm()
UserForm1.Show
End Sub

Public Sub eICount()
Call ExportAndImportCounter
End Sub


Public Sub ExportAndImportCounter()
Dim wbNum As Integer 'カウントするワークブックの数
Dim wsNum As Integer 'カウントするワークシートの数
Dim i As Integer, l As Integer 'ループの変数
Dim pICount As Integer '担当者数
Dim monthCount As Integer '月をまたぎの差
Dim startDate As Date, endDate As Date '検索期間
Dim countBook(3) As String '件数取得ブックのパス
Dim findSheetName As Range '検索するシート名
Dim yesNo As String
Dim getTextBook As Workbook 'マクロのあるシート
Dim hwnd As LongPtr
Dim beforeHwnd As LongPtr
Dim bookPath As Variant, bookName As Variant, pIName() As Variant
Dim pIChange As Range, iRange As Range

'On Error GoTo err

Application.ScreenUpdating = False
Erase exImCount '配列を初期値にする
Set getTextBook = ActiveWorkbook

startDate = InputBox("検索開始日を指定してください。" & vbCrLf & "過去1週間分が対象になります。", "検索開始日入力画面", Date - 6)
endDate = startDate + 6

monthCount = Mid(endDate, 6, 2) - Mid(startDate, 6, 2)

For l = 0 To monthCount '月をまたぐ場合ループする

    With getTextBook.Worksheets("輸出入件数取得シート")
    
        countBook(0) = .Range("D" & 6 + 5 * l).Text '輸入の台帳
        countBook(1) = .Range("D" & 7 + 5 * l).Text '2階関係の台帳
        countBook(2) = .Range("D" & 8 + 5 * l).Text '輸出の台帳
        countBook(3) = .Range("D9").Text '転記対象のブック
        
        '担当者ごとの件数取得--------
        'Set pIChange = .Cells.Find(what:="担当者名", lookat:=xlWhole).Offset(1)
        'Set pIChange = .Range(Cells(pIChange.Row, pIChange.Column), Cells(pIChange.End(xlDown).Row, pIChange.Column))
        'For Each iRange In pIChange '担当者
            
           ' If iRange.Offset(, 3).Value = "○" Then
                
               'pIName(pICount, 0) = iRange.Value
            
            'End If
            
            'pICount = pICount + 1
            
        'Next
         '----------------------------
          
    End With
    
    For wbNum = 0 To 2 'ブックの数だけループ
    
        bookName = Split(countBook(wbNum), "\")
        hwnd = FindWindow("XLMAIN", bookName(UBound(bookName)) & " - Excel")
        
        If hwnd = 0& Then
            Workbooks.Open countBook(wbNum), ReadOnly:=True
        Else
            beforeHwnd = hwnd
            SetForegroundWindow beforeHwnd
        End If
        
        DoEvents
        
        With ActiveWorkbook
            wsNum = .Worksheets.Count
            For i = 1 To wsNum
            'MsgBox Worksheets(i).Name
            
            'シート名がリストにない場合-------------------
            Set findSheetName = getTextBook.Worksheets("輸出入件数取得シート").Columns(1).Find(what:=.Worksheets(i).Name, lookat:=xlWhole)
            If findSheetName Is Nothing Then
                ThisWorkbook.Worksheets("輸出入件数取得シート").Range("A1").End(xlDown).Offset(1, 0).Value = .Worksheets(i).Name
                yesNo = MsgBox(.Worksheets(i).Name & "は見つかりません。" & vbCrLf & "リストに追加しますか？", vbYesNo)
                If yesNo = vbYes Then
                    ThisWorkbook.Worksheets("輸出入件数取得シート").Range("A1").End(xlDown).Offset(, 1).Value = "○"
                Else
                    ThisWorkbook.Worksheets("輸出入件数取得シート").Range("A1").End(xlDown).Offset(, 1).Value = "×"
                End If
                
            End If
            '-------------------
            
            
                '取得設定が"○"であり、非表示でないもの　！findのafterを設定！
                If getTextBook.Worksheets("輸出入件数取得シート") _
                .Columns(1).Find(what:=.Worksheets(i).Name, lookat:=xlWhole) _
                .Offset(, 1).Value = "○" And .Worksheets(i).Visible = True Then
                    Call getCount(startDate, endDate, Worksheets(i).Name, wbNum)
                End If
                
            Next
            
        End With
        
        ActiveWorkbook.Close (False)
        DoEvents
        
    Next

Next

MsgBox startDate & "～" & endDate & vbCrLf & "の輸入件数は" & _
exImCount(0) & "件" & vbCrLf & "輸入2階関係は" & exImCount(1) & "件" & vbCrLf & "輸出件数は" & exImCount(2) & "件" & vbCrLf & "輸出2階関係は" & exImCount(3) & "件です。"

With getTextBook.Worksheets("輸出入件数取得シート")
    .Range("D1").Value = startDate
    .Range("E1").Value = endDate
    .Range("D2").Value = exImCount(0)
    .Range("E2").Value = exImCount(1)
    .Range("D3").Value = exImCount(2)
    .Range("E3").Value = exImCount(3)
End With

'転記するブックをアクティブにする
bookName = Split(countBook(3), "\")
hwnd = FindWindow("XLMAIN", bookName(UBound(bookName)) & " - Excel")

If hwnd = 0& Then
    Workbooks.Open countBook(3)
Else
    beforeHwnd = hwnd
    SetForegroundWindow beforeHwnd
End If

Call tenki(startDate, endDate) '転記処理

Application.ScreenUpdating = True

Exit Sub

'エラー処理-start-
err:
MsgBox "エラーが発生したため、処理を中断しました。" & vbCrLf & err.Number & ":" & err.Description
Application.ScreenUpdating = True
'エラー処理-end-

End Sub

Public Sub getCount(ByVal startDate As Date, ByVal endDate As Date, ByVal sheetName As String, ByVal bookCount As Integer)
Dim numStartRow As Integer, numEndRow As Integer '検索対象の最終行
Dim numCol As Integer '"番号"のある列
Dim permissionCol As Integer '"許可日"のある列
Dim perCount As Integer '許可日のカウント
Dim permissionCells As Range '許可入力されているセル範囲
Dim i As Range

With Worksheets(sheetName)

    .Activate
    numStartRow = .Cells.Find(what:="番号", lookat:=xlWhole).Row + 1
    numCol = .Cells.Find(what:="番号", lookat:=xlWhole).Column
    numEndRow = .Cells(numStartRow, numCol).End(xlDown).Row
    permissionCol = .Cells.Find(what:="許可日", lookat:=xlWhole).Column
    Set permissionCells = .Range(Cells(numStartRow, permissionCol), Cells(numEndRow, permissionCol))
    
        For Each i In permissionCells
            '許可日が空でなく、検索開始日以上、検索終了日以下ならカウントする。
            If i.Value <> "" And i >= startDate And i <= endDate Then
                i.Interior.ColorIndex = 3 '色付けして確認したいとき。
                perCount = perCount + 1
            End If
            
        Next
        
End With

If ActiveSheet.Name = "輸出" Then bookCount = 3
exImCount(bookCount) = exImCount(bookCount) + perCount

End Sub

Public Sub tenki(ByVal startDate As Date, ByVal endDate As Date)
Dim selectCol1 As Integer, selectcol2 As Integer
Dim startRow As Integer, endRow1 As Integer, endrow2 As Integer
Dim finalRow As Integer
Dim selectCell1 As Range, selectCell2 As Range, i As Range, dateCell As Range

With ActiveWorkbook.ActiveSheet 'シート名変えられるように

    startRow = .Cells.Find(what:="期間", lookat:=xlWhole).Row + 1
    selectCol1 = .Cells.Find(what:="期間", lookat:=xlWhole).Column
    selectcol2 = .Cells.Find(what:="期間", lookat:=xlWhole, after:=.Cells.Find(what:="期間")).Column
    endRow1 = .Cells(Rows.Count, selectCol1).End(xlUp).Row
    endrow2 = .Cells(Rows.Count, selectcol2).End(xlUp).Row
    finalRow = 0
    
    Set selectCell1 = .Range(Cells(startRow, selectCol1), Cells(endRow1, selectCol1))
    Set selectCell2 = .Range(Cells(startRow, selectcol2), Cells(endrow2, selectcol2))
    
    '新年度のシートを自動作成-----
    Call createSheet(endrow2, selectcol2)
    '-----------------------------
    
    For Each i In selectCell1 '1つ目の表を検索
    
        If i.Value >= startDate And i <= endDate Then
        
            finalRow = i.Row
            i.Offset(, 3).Value = exImCount(0) + exImCount(1)
            i.Offset(, 4).Value = exImCount(2) + exImCount(3)
            
        End If
        
    Next
    
    If finalRow = 0 Then
    
        For Each i In selectCell2 '2つ目の表を検索
        i.Select
            If i.Value >= startDate And i <= endDate Then
            
                finalRow = i.Row
                i.Offset(, 3).Value = exImCount(0) + exImCount(1)
                i.Offset(, 4).Value = exImCount(2) + exImCount(3)
                
            End If
            
        Next

    End If
    
    
End With

End Sub

Public Sub createSheet(ByVal endrow2 As Integer, ByVal selectcol2 As Integer) '新年度のシート作成
Dim beforeDate As Date

With ActiveWorkbook

    If .Worksheets(Worksheets.Count).Cells(endrow2, selectcol2).Offset(0, 4).Value <> "" Then
        
        beforeDate = .Worksheets(Worksheets.Count).Cells(endrow2, selectcol2).Value
        .Worksheets("原紙").Copy after:=Worksheets(Worksheets.Count)
        .Worksheets(Worksheets.Count).Name = ThisWorkbook.Worksheets("輸出入件数取得シート").Range("D14").Text
        .Worksheets(Worksheets.Count).Cells.Find(what:="期間", lookat:=xlWhole).Offset(1, 0).Value = beforeDate + 1
        
    End If
    
End With

End Sub

Public Sub personInCharge()
Dim inCharge() As String
Dim perRow As Integer, perCol As Integer
Dim perRange As Range, i As Range
Dim personsName As Variant

perRow = ThisWorkbook.Worksheets("担当者名簿").Rows(1).Find(what:="苗字", lookat:=xlWhole).Row
perRow = ThisWorkbook.Worksheets("担当者名簿").Rows(1).Find(what:="苗字", lookat:=xlWhole).Column
perRange = ThisWorkbook.Worksheets("担当者名簿").Range(Cells(perRow, perCol), Cells(Rows.Count, perCol))

For Each i In perRange

    'redim preserve

Next


End Sub


