Attribute VB_Name = "Module4"
Sub 転記(ByVal strNewWorkbookName As String, TABnumber As Integer)
Dim TENKIBOOK As Workbook
Dim newBookname, newSheetname, newSelectSheet As String
Dim WORKSCON, MsgSelect As Integer
Dim TORF As Boolean
    TORF = False
    
    If OpenedBook(TABnumber) = 0 Then '転記先ファイルが開いていなかったら開く
        If CreateObject("SCRIPTING.FILESYSTEMOBJECT"). _
        FILEEXISTS((Workbooks(WWW).Worksheets("設定").Cells(3 + TABnumber, 5).Value)) = True Then 'ファイルが無かった場合の処理を行う
        Workbooks.Open (Workbooks(WWW).Worksheets("設定").Cells(3 + TABnumber, 5).Value)
        Else
        MsgSelect = MsgBox("指定されたファイルが見つかりません。" & vbCrLf & "ファイルを直接指定するか、オプションから設定しなおしてください。", vbYesNo, vbExclamation)
            If MsgSelect = vbYes Then
            newSelectSheet = Application.GetOpenFilename("*,*.xlsx", Title:="転記先のファイルを指定してください。")
                If newSelectSheet <> "False" Then
                Workbooks.Open (newSelectSheet)
                Else
                MsgBox ("キャンセルされました。作業を中断します。"), vbExclamation
                Exit Sub
                End If
            Else
            MsgBox ("キャンセルされました。作業を中断します。"), vbExclamation
            Exit Sub
            End If
        End If
    Else
    Workbooks(Dir(Workbooks(WWW).Worksheets("設定").Cells(3 + TABnumber, 5).Value)).Activate
    End If
    
Set TENKIBOOK = ActiveWorkbook
    For Each WS In TENKIBOOK.Worksheets '同名のシートがないか調べ、ある場合は追加。無い場合は新規シートを作成する。
        If WS.Name = Mid(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="納入指定日1").Offset(1, 0).Value, 5, 2) _
        & "." & Right(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="納入指定日1").Offset(1, 0).Value, 2) Then
        TORF = True
        newSheetname = WS.Name
        End If
    Next
        If TORF = True Then
        TENKIBOOK.Worksheets(newSheetname).Activate
        Else
        TENKIBOOK.Activate
        WORKSCON = TENKIBOOK.Worksheets.Count
        TENKIBOOK.Worksheets("原紙").Copy AFTER:=TENKIBOOK.Worksheets(WORKSCON - 1)
        End If
    
    

Call HHHDD(WORKSCON, TORF, strNewWorkbookName, TENKIBOOK)
End Sub


Sub HHHDD(ByVal WORKSCON As Integer, TORF As Boolean, strNewWorkbookName As String, TENKIBOOK As Workbook)

'strNewWorkbookName=指示書 TENKIBOOK=作業用ブック


'!追加注文の場合備考欄に追加と付ける!
Dim TAR, nextTAR, DAYVVAL As String
Dim R, TARcol, TARrow As Integer
R = 0
TARrow = TENKIBOOK.ActiveSheet.Cells(Rows.Count, TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="部品番号").Column).End(xlUp).Offset(1, 0).Row
Workbooks(strNewWorkbookName).Activate
For Each RANGVAL In Range("A1:S1") '!TENKIBOOKにうつす時のTARrow,TARcolの後のENDとOFFSETいるのか検証!
    TAR = RANGVAL.Text
    
    If TAR = "発注者品名ｺｰﾄﾞ-備考" Or TAR = "発注者品名ｺｰﾄﾞ-納入時刻2" Then '黒田の場合後者になる
        nextTAR = "部品番号"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        R = TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).End(xlUp).Offset(1, 0).Row
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "注文番号" Then '!記入がない場合上に詰めてしまので直す!
        nextTAR = "P/O　No."
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "納入指示数量1" Then
        nextTAR = "出庫数量"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="LOT No.").Offset(-1, 0).Value = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="出庫数量").Offset(-1, 0).Value

    
    ElseIf TAR = "受渡場所名" Then
        nextTAR = "納入場所"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "品名(品名仕様)" Then '生方のみ品名記入欄がある為
        If Workbooks(strNewWorkbookName).Worksheets(1).PageSetup.LeftHeader Like "*生方*" Then
        nextTAR = "部品名"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        End If
    ElseIf TAR = "機種ｺｰﾄﾞ" Then '記入がない場合飛ばす
        nextTAR = "P/O　No."
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Offset(1, 2).Column
        If Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp).Value <> RANGVAL.Value Then
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp).Offset(0, 1)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        End If
    ElseIf TAR = "納入指定日1" Then '部品番号の行番号を取得し、納品日付に記入する
      nextTAR = "納品日付"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        For Each LTAR In Range(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(1, 0), Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp))
            If TENKIBOOK.ActiveSheet.Cells(R, TARcol + 1).Value <> "" Then
            TENKIBOOK.ActiveSheet.Cells(R, TARcol).Value = Mid(LTAR.Value, 5, 2) & "/" & Right(LTAR.Value, 2) '日付入力
            End If
            R = R + 1
        Next
        '作業用シートによっては、1行上のライン貼り付け箇所を考える
    End If
Next



DAYVVAL = Mid(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="納入指定日1").Offset(1, 0).Value, 5, 2) _
& "." & Right(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="納入指定日1").Offset(1, 0).Value, 2)
If TORF = False Then
TENKIBOOK.Worksheets(WORKSCON).Name = DAYVVAL
End If
TENKIBOOK.Save
End Sub

Sub KLFJDFK()
TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="出庫数量").Offset(-1, 0).Value = _
WorksheetFunction.Sum(TENKIBOOK.ActiveSheet.Range(TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="出庫数量").Offset(1, 0), Cells(Rows.Count, TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="出庫数量").Column).End(xlUp)))
End Sub
