Attribute VB_Name = "Module1"
Option Explicit '2015/08/23作成
'新規ワークブック名
Dim strNewWorkbookName As String  '転記先 指示書
Dim strCsvbookName, NNN As String '転記元 CSV
Dim lngRowCounter As Long
Dim lngCellCounter As Long
Dim SAVRAN As Range
Public KA, KC, TABnumber As Integer
Public lngOpenBookNumber As Long
Public SPNAME, WWW, addNo As String
Public SW As Byte
Public WB1, WB2, WB3 As Workbook

Function f_NewWorkbook() As String

    '新しいBookオープン
    Workbooks.Add
    f_NewWorkbook = ActiveWorkbook.Name
    
End Function

Public Sub Main(ByVal valListIndex As Variant, CATE As String)
SW = 0
NNN = ""
addNo = "1"

    If valListIndex = -1 Then
        MsgBox "リストから選択してください"
        Exit Sub
    End If

    'ファイルオープン
    If valListIndex = 3 Then
    Call 黒田複数
    Else
    Call OpenFile
    End If
    'CSVファイル名のセット
    strNewWorkbookName = f_NewWorkbook
    
    lngOpenBookNumber = Workbooks.Count
    
    '正和シール
    If valListIndex = 0 Or valListIndex = 2 Then
        Call 注文番号
        Call 部品名
        Call 発注者品名
        Call 納入指示数
        Call 納入指定日
'''20141111追加
        Call 納入時刻
        
        Call 受渡場所
        Call 市町村
        Call 設計変更番号
        Call 機種コード
        Call 発注部門名
        
        Call 並び替え正和
        '20160111
        Call 削除(valListIndex)
    'SKK
   ElseIf valListIndex = 1 Then
        Call 注文番号
        Call 部品名
        Call 発注者品名
        Call 納入指示数
        Call 納入指定日
'''20141111追加
        Call 納入時刻
        
        Call 受渡場所
        Call 市町村
        Call 設計変更番号
        Call 機種コード
        Call 発注部門名
        
       
        '20160111
        Call 削除(valListIndex)
        '20161104黒田追加
    ElseIf valListIndex = 3 Or valListIndex = 4 Then
        Call 注文番号
        Call 部品名
        Call 発注者品名
        Call 納入指示数
        Call 納入指定日
        Call 納入時刻
        Call KURODAUKEWATASHI
        
    
    
    
        
    '選択せず
    Else
        MsgBox "リストから選択してください"
        Exit Sub
    End If
    
    '追加部分
    Dim AAAF As Integer
    
    If UserForm1.ComboBox1.Text = "正和シール" Or UserForm1.ComboBox1.Text = "SKK" Then
    AAAF = 28
    ActiveSheet.PageSetup.PaperSize = xlPaperB4
    ActiveWindow.Zoom = 70
   
    Call KDODNA
    
    ElseIf UserForm1.ComboBox1.Text = "SKK" Then
    AAAF = 25
    ActiveSheet.PageSetup.PaperSize = xlPaperB4
    ActiveWindow.Zoom = 70
    Else
    AAAF = 35
    
    End If
    
    Cells.Select
    Selection.RowHeight = AAAF
        
        'ｻﾌﾟﾗｲﾔｰ名取得
        Dim FIFI As Range
        Dim HINBANUP As String
        Dim FCOL As Integer
        Dim FROW As Integer
        Dim VBS As Integer
        Dim ACS As Worksheet
        
        
        Set ACS = ActiveWorkbook.ActiveSheet
        If valListIndex = 3 Or valListIndex = 4 Then  '黒田の場合品番１つ上の項目名が違う為
        HINBANUP = "発注者品名ｺｰﾄﾞ-納入時刻2"
        Else
        HINBANUP = "発注者品名ｺｰﾄﾞ-備考"
        End If
        Set FIFI = Workbooks(WWW).Worksheets("DATABASE").Cells.Find _
        (WHAT:=ACS.Cells.Find(HINBANUP).End(xlDown).Value, LookIn:=xlValues, LOOKAT:=xlWhole)
            If FIFI Is Nothing Then
            VBS = MsgBox(ACS.Cells.Find(HINBANUP).End(xlDown).Value & "が見つかりませんでした。品番を追加してから再始動するか、" & vbCrLf & "ｻﾌﾟﾗｲﾔｰ名を記入してください。", vbOKCancel)
                 
                If VBS = vbOK Then
                UserForm2.Show (vbModal)
                If SW = 3 Then 'フォームのキャンセルボタンが押されたら終了する。
                Workbooks(strCsvbookName).Close
                GoTo ErrorHandler
                End If
                GoTo SNA
                ElseIf VBS = vbCancel Then
                Workbooks(strCsvbookName).Close
                GoTo ErrorHandler
                End If
            End If
        
        FCOL = FIFI.Column
        FROW = FIFI.Row
        SPNAME = Workbooks(WWW).Worksheets("DATABASE").Cells(FROW, FCOL).End(xlUp).Value
        TABnumber = TABno(SPNAME)
SNA:
        TABnumber = TABno(SPNAME)

        'ヘッダー、フッター、行高等を調整
        With ActiveSheet.PageSetup
        Dim NDAY As String
        NDAY = ACS.Cells.Find("納入指定日1").Offset(1, 0).Value
        .Orientation = xlLandscape
        If valListIndex <> 3 And valListIndex <> 4 Then
        .LeftHeader = "&13 " & "&B" & Mid(NDAY, 5, 2) & "/" & Mid(NDAY, 7, 2) & "  " & SPNAME
        Else
        .LeftHeader = "&13 " & "&B" & Mid(NDAY, 5, 2) & "/" & Mid(NDAY, 7, 2) & "  " & SPNAME & vbCr & KA & "/" & KC & "件"
        End If
        .RightHeader = "&B" & "&P" & "/" & "&N"
        End With
        Columns("A:A").ColumnWidth = 6
        Columns(ACS.Cells.Find("注文番号").Column).AutoFit '品名、品番、受渡し場所等の列を調整するようにマクロかく2016/09/05
        Columns(ACS.Cells.Find("納入指定日1").Column).ColumnWidth = 8.5 '数量
        Columns(ACS.Cells.Find("品名(品名仕様)").Column).ColumnWidth = 6 '品名
        Columns(ACS.Cells.Find("受渡場所名").Column).ColumnWidth = 6
        Columns(ACS.Cells.Find("納入指定日1").Column).ColumnWidth = 11.5
        Columns(ACS.Cells.Find("納入指示数量1").Column).ColumnWidth = 7.5
        Columns(ACS.Cells.Find(HINBANUP).Column).ColumnWidth = 15
        Columns("G:K").ColumnWidth = 5.5 '受渡場所以降
        
        Select Case SPNAME
        Case "昭和機器工業"
            Call SKKSort
            Call KEISEN(SPNAME)
        
            Columns(ACS.Cells.Find("発注者品名ｺｰﾄﾞ-備考").Column).ColumnWidth = 20
            Columns(ACS.Cells.Find("受渡場所名").Column).ColumnWidth = 12
            Columns("A:A").ColumnWidth = 3

            Call SKKEREMA
        Case "正和シール販売"
        If UserForm1.ComboBox1.Text = "正和シール" Then
            Columns(ACS.Cells.Find("納入時刻1").Column).Delete
        End If
            Call KEISEN(SPNAME)
        
        Case "三和　桐生工場"
            Call KIRYU
            Call KEISEN(SPNAME)
        Case "生方製作所"
            Call KEISEN(SPNAME)
        Case "三和　前橋工場"
            Call KEISEN(SPNAME)
            Call MAEBASI
        Case "太平洋工業"
            Call KEISEN(SPNAME)
        Case "エレマテック"
            Call KEISEN(SPNAME)
        Case "黒田製作所"
            Cells.RowHeight = 20
            Call KURODAKEISEN
            
        End Select
        
    '参照元CSVファイルのクローズ
    Workbooks(strCsvbookName).Close SaveCHANGES:=False
        
        If Dir(Workbooks(WWW).Worksheets("設定").Cells(3 + TABnumber, 4).Value, vbDirectory) = "" Then '指定されたパスが無い場合はデスクトップに保存。
        MsgBox ("指定されたフォルダが見つかりませんでした。" & vbCrLf & "デスクトップに保存します。"), vbInformation
        Set SAVRAN = Workbooks(WWW).Worksheets("設定").Range("D100")
        Else
        Set SAVRAN = Workbooks(WWW).Worksheets("設定").Cells(3 + TABnumber, 4)
        End If
        
        
        Dim SheetNo As String
        SheetNo = (Left(NDAY, 4) & Mid(NDAY, 5, 2) & Mid(NDAY, 7, 2) & Trim(SPNAME) & "様納入分指示書" & NNN & ".xlsx")
            If CATE <> "" Then
            
            NNN = CATE & addNo '項目を追加
                
                Do While DIRECT(SAVRAN, NDAY, NNN) = 1 '同名のシートが無いか検索
                addNo = addNo + 1 '同名のシートがあったら名前を変えて再検索
                NNN = CATE & addNo
                SheetNo = (Left(NDAY, 4) & "." & Mid(NDAY, 5, 2) & "." & Mid(NDAY, 7, 2) & Trim(SPNAME) & "様納入分指示書" & NNN & ".xlsx")
                Loop
            SheetNo = (Left(NDAY, 4) & Mid(NDAY, 5, 2) & Mid(NDAY, 7, 2) & Trim(SPNAME) & "様納入分指示書" & NNN & ".xlsx")
   
            End If
       

       
    Workbooks(strNewWorkbookName).SaveAs Filename:=SAVRAN.Value & "\" & SheetNo
    strNewWorkbookName = ActiveWorkbook.Name
    Call 転記(strNewWorkbookName, TABnumber)
    Workbooks(WWW).Save
    Application.CutCopyMode = False

    MsgBox "正常に終了しました"
Exit Sub
ErrorHandler:
MsgBox ("キャンセルされました。")
Workbooks(strNewWorkbookName).Close SaveCHANGES:=False
End Sub

Private Sub 削除(valListIndex As Variant)
    Dim lngLastRow As Long

    With Workbooks(lngOpenBookNumber).Worksheets(1)
        '最終行取得
        lngLastRow = .Range("B65536").End(xlUp).Row
       
        .Activate
        
        If valListIndex = 0 Then
           
            
            'セルサイズの調節

            .Columns("B:L").EntireColumn.AutoFit
        
        ElseIf valListIndex = 2 Then
            'J設計変更
            .Range(Cells(1, 10), Cells(lngLastRow, 10)).Delete Shift:=xlShiftToLeft
            
             'I市町村コード
            .Range(Cells(1, 9), Cells(lngLastRow, 9)).Delete Shift:=xlShiftToLeft
            
            'G時刻
            .Range(Cells(1, 7), Cells(lngLastRow, 7)).Delete Shift:=xlShiftToLeft
            'セルサイズの調節
            .Columns("B:I").EntireColumn.AutoFit
       End If
    End With

End Sub

Private Sub OpenFile()

'開くCSVファイルの指定を行う
On Error GoTo ErrorHandler
    strCsvbookName = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")
    Workbooks.Open strCsvbookName
    strCsvbookName = Dir(strCsvbookName)
Exit Sub

'キャンセルでもこっちの処理
ErrorHandler:
        'MsgBox "選択しなかったので終了します"
        Exit Sub
End Sub

Private Sub 注文番号()

    Dim A As Long
    
    A = 1
    lngCellCounter = 1
    
    Do While Workbooks(strCsvbookName).Worksheets(1).Range("D" & A).Value <> ""
        Workbooks(strCsvbookName).Worksheets(1).Range("D" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
        A = A + 1
    Loop
    
    '繰り返し行数確定
    lngRowCounter = A
    
End Sub

Private Sub 部品名()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("H" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 発注者品名()
    Dim A As Long
    Dim strTMP As String
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
'    For a = 1 To lngRowCounter
'        Workbooks(strCsvbookName).Worksheets(1).Range("J" & a).Copy Destination:= _
'        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(a, lngCellCounter)
'    Next

'20140414 備考欄にランクが入っていたら、「-○」と発注者品名の後ろに付ける

    For A = 1 To lngRowCounter
        strTMP = Workbooks(strCsvbookName).Worksheets(1).Range("J" & A).Value
        
        '備考欄が空白ではなかったら
        If Workbooks(strCsvbookName).Worksheets(1).Range("U" & A).Value <> "" Then
            strTMP = strTMP & "-" & Workbooks(strCsvbookName).Worksheets(1).Range("U" & A).Value
        End If
        
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter) = strTMP
        
    Next


End Sub

Private Sub 納入指示数()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("L" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 納入指定日()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("M" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 納入時刻()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("P" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 受渡場所()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("V" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 市町村()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AE" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 設計変更番号()
    Dim A As Long
       
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AF" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 機種コード()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AG" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub 発注部門名()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AI" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub KURODAUKEWATASHI()
    Dim A As Long
    
    '隣の列へ移動
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("FP" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next
 KA = Range(Cells(2, 3), Cells(Rows.Count, 3).End(xlUp)).Count
 Workbooks(strNewWorkbookName).Worksheets("Sheet1").Rows(1).AutoFilter FIELD:=7, Criteria1:="ｲｼｻﾞｶｸﾐﾀﾃ"
 Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells.Sort KEY1:=Worksheets("Sheet1").Range("C1"), _
 ORDER1:=xlAscending, Header:=xlYes
 KC = Range(Cells(2, 3), Cells(Rows.Count, 3).End(xlUp)).SpecialCells(xlCellTypeVisible).Count
End Sub
Private Sub 並び替え正和()
    Dim lngLastRow As Long
    Dim i As Integer
    Dim i2 As Integer
    Dim strTmpText As String
    Dim lngTextNumber As Long
    Dim lngTmpRow  As Long
    Dim blnSpaceFlag As Boolean
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        
        '空白行フラグを初期化
        blnSpaceFlag = False
       
        '最終行取得
        lngLastRow = .Range("A65536").End(xlUp).Row
       
        .Activate
       
        'C2発注者品名ｺｰﾄﾞ　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
              
        '最終行まで回す
        For i2 = 2 To lngLastRow
            'G2以下　受渡場所名　ｻﾝﾜﾃｯｸ　を　最終行+空白行の下へ
            strTmpText = .Cells(i2, 7).Text
            '含まれていたら、最終行+空白行の下へ
            If InStr(strTmpText, "ｻﾝﾜﾃｯｸ") Or InStr(strTmpText, "ｻﾝﾜ ｵｵﾀ1 ｳｹｿ") > 0 Then
                '空白行が入ってなかったら入れる
                If blnSpaceFlag = False Then
                   blnSpaceFlag = True
                   lngLastRow = lngLastRow + 1
                End If
                
                '行をカットして最終行の下にペースト
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '空白まで来たら抜ける
            If strTmpText = "" Then
                Exit For
            End If
        Next
    
        '空白行フラグを初期化
        blnSpaceFlag = False
       
        '最終行まで回す
        For i2 = 2 To lngLastRow
           
           '20160111 ﾀｲﾍｲﾖｳｺｳｷﾞｮｳも追加
            'G2以下　受渡場所名　ｾｲｺｰﾚｼﾞﾝ　を　最終行+空白行の下へ
            strTmpText = .Cells(i2, 7).Text
            '含まれていたら、最終行+空白行の下へ
            If InStr(strTmpText, "ｾｲｺｰﾚｼﾞﾝ") > 0 Or InStr(strTmpText, "ﾀｲﾍｲﾖｳｺｳｷﾞ") > 0 Then
                '空白行が入ってなかったら入れる
                If blnSpaceFlag = False Then
                    blnSpaceFlag = True
                    lngLastRow = lngLastRow + 1
                End If
               
                '行をカットして最終行の下にペースト
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '空白まで来たら抜ける
            If strTmpText = "" Then
                Exit For
            End If
        Next
       
        '空白行フラグを初期化
        blnSpaceFlag = False
        
        For i2 = 2 To lngLastRow
            '発注者品名コードが
            strTmpText = Trim(.Cells(i2, 3).Text)
            lngTextNumber = Len(strTmpText)
               
            '5文字 & "-" & 5文字の場合で
            If lngTextNumber = 11 Then
            
                '10文字目が"-"だったら何もしない
                If InStr(10, strTmpText, "-") <> 10 Then
                
                    strTmpText = Left(strTmpText, 1)
                       
                    '頭の文字が"5"じゃなかった場合
                    If strTmpText <> "5" Then
                        'その行削除
                        .Rows(i2).Delete
                        lngLastRow = lngLastRow - 1
                        i2 = i2 - 1
                    Else
                        '空白行が入ってなかったら入れる
                        If blnSpaceFlag = False Then
                            blnSpaceFlag = True
                            lngLastRow = lngLastRow + 1
                        End If
                        
                        '行をカットして最終行の下にペースト
                        .Rows(i2).Cut
                        .Rows(lngLastRow + 1).Insert
                        i2 = i2 - 1
                       
                    End If
                End If
            End If
                    
            '空白まで来たら抜ける
            If strTmpText = "" Then
                Exit For
            End If
        Next
           
        '空白行フラグを初期化
        blnSpaceFlag = False
        
        '最終行まで回す
        For i2 = 2 To lngLastRow
           
            'G2以下　受渡場所名　CKD　を　最終行+空白行の下へ
            strTmpText = .Cells(i2, 7).Text
            '含まれていたら、最終行+空白行の下へ
            If InStr(strTmpText, "CKD") > 0 Then
                '空白行が入ってなかったら入れる
                If blnSpaceFlag = False Then
                    blnSpaceFlag = True
                    lngLastRow = lngLastRow + 1
                End If
               
                '行をカットして最終行の下にペースト
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '空白まで来たら抜ける
            If strTmpText = "" Then
                Exit For
            End If
            
        Next
        
        
        '上から空白行が出てくるまでの間を、F2以下　納入時刻1　昇順で並び替え
        lngTmpRow = .Range("A1").End(xlDown).Row
        
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 6), ORDER1:=xlAscending
        
        'A2注文番号　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 1), ORDER1:=xlAscending
        
        'C2発注者品名ｺｰﾄﾞ　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 3), ORDER1:=xlAscending
        
        'A列に1列挿入
        .Columns(1).Insert
            
        'セルサイズの調節
        .Columns("B:L").EntireColumn.AutoFit
        
    End With
End Sub

Private Sub 並び替えSKK()
    Dim lngLastRow As Long
    Dim i As Integer
    Dim strTmpText As String
    Dim blnSpaceFlag As Boolean
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        
        '空白行フラグを初期化
        blnSpaceFlag = False
       
        '最終行取得
        lngLastRow = .Range("A65536").End(xlUp).Row
       
        .Activate
       
        'C2発注者品名ｺｰﾄﾞ　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
                
        'J2機種ｺｰﾄﾞ　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 10), ORDER1:=xlAscending
                
        '最終行まで回す
        For i = 2 To lngLastRow
            '発注部門名が空だったら削除
            strTmpText = .Cells(i, 11).Text
            If strTmpText <> "" Then
                blnSpaceFlag = True
            End If
        Next
        
        If blnSpaceFlag = False Then
            .Columns(11).Delete
        End If
        
        'セルサイズの調節
        .Columns("A:K").EntireColumn.AutoFit
                
    End With
    
End Sub

Private Sub 並び替えその他()
    Dim lngLastRow As Long
    Dim i2 As Integer
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        '最終行取得
        lngLastRow = .Range("A65536").End(xlUp).Row
        
        'C2発注者品名ｺｰﾄﾞ　昇順で並び替え
        .Range(Cells(2, 1), Cells(lngLastRow, 8)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
             
                
        'セルサイズの調節
        .Columns("A:H").EntireColumn.AutoFit
    End With
   
End Sub


Private Sub KEISEN(ByVal SPNAME As String) 'ライン毎に罫線を引く
Dim AHA As Range
Set AHA = Cells(Rows.Count, 4).End(xlUp)

Dim AST, STR, LineSt As Integer
LineSt = xlContinuous
AST = ActiveSheet.Cells.Find("機種ｺｰﾄﾞ").Offset(1, 0).Column
    If UserForm1.ComboBox1.Text = "SKK" Then
    STR = 0
    ElseIf UserForm1.ComboBox1.Text = "正和シール" Then
    STR = 1
    Else
    STR = -3
    End If
Dim PPPC As Range
    For Each PPPC In Range(Cells(2, AST), Cells(AHA.Row, AST))
        If PPPC.Row = AHA.Row Then Exit Sub
        If PPPC.Offset(1, -6).Value = "" Then
        GoTo NNN
        End If
        If PPPC.Value = PPPC.Offset(1, 0).Value Then 'ﾗｲﾝが変わらなければ薄い罫線を引く
           
        Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).ColorIndex = 15
       
        GoTo NNN
        Else
            If PPPC.Offset(1, 0).Value <> PPPC.Value Then 'ﾗｲﾝが変わったら罫線を引く
            
            Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).ColorIndex = 56
            
            End If
        End If
    
NNN:
    
    Next PPPC
End Sub

Sub KDODNA() '改ページ

Dim KRANGE As Range
Set KRANGE = Range("D1").End(xlDown).Offset(2, 0)
ActiveSheet.HPageBreaks.Add BEFORE:=KRANGE
End Sub



Sub 黒田複数() '!CSVの数だけ処理を行う!
Dim A, D, E As Variant
Dim WC, R, C, P As Integer
Dim KURODACSV As New Collection
WC = 0
E = 1

A = Application.GetOpenFilename("CSVファイル(*.csv),*.csv", Title:="Ctrlを押しながらcsvﾃﾞｰﾀを2つ以上選択してください。", MultiSelect:=True)
    
    For Each D In A
    Workbooks.Open D
    KURODACSV.Add Item:=ActiveWorkbook
         WC = WC + 1
    Next
    
    Do While E <= UBound(A) 'ワークブックの数だけ繰り返す。
        If E <> 1 Then
            KURODACSV(E).Activate
            If KURODACSV(E).Worksheets(1).Range("B2").Offset(1, 0).Value = "" Then
            KURODACSV(E).Worksheets(1).Range(Cells(2, 1), Cells(2, Columns.Count).End(xlToLeft)).Copy
            Else
            R = KURODACSV(E).Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row '最終行
            C = KURODACSV(E).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '最終列
            KURODACSV(E).Worksheets(1).Range(Cells(2, 1), Cells(R, C)).Copy
            End If
        Else
        GoTo SSS
        End If
        
        KURODACSV(1).Worksheets(1).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial
        Application.CutCopyMode = False
        KURODACSV(E).Close SaveCHANGES:=False

SSS:
        E = E + 1
    Loop
    strCsvbookName = KURODACSV(1).Name
    
    
End Sub

Sub KURODAKEISEN()
Dim KURO As Range
Dim KKR, KKC, LASTKR As Integer
LASTKR = Cells(Rows.Count, 1).End(xlUp).Row
    For Each KURO In Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp))
        If KURO.Row = LASTKR Then
        Exit Sub
        End If
    KKR = KURO.EntireRow.Row
    KKC = KURO.EntireColumn.Column
    Range(Cells(KKR, KKC), Cells(KKR, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(KKR, KKC), Cells(KKR, 15)).Borders(xlEdgeBottom).ColorIndex = 15
    Next
End Sub


