Attribute VB_Name = "Module3"
Sub MAEBASI()
Dim MAEOG As Range
Dim A As Integer

For Each MAEOG In Range(Cells(Cells.Find(WHAT:="発注者品名ｺｰﾄﾞ-備考").Offset(1, 0).Row, 4), Cells(Rows.Count, 4).End(xlUp))
MAEOG.Select
If MAEOG.Value Like "*-0290" Or MAEOG.Value Like "*-0291" Or MAEOG.Offset(0, 4).Value = "KF" Then '文字列に-0290,又は-0291が含まれていた場合の判定
MAEOG.EntireRow.Font.Strikethrough = True
End If
Next
End Sub

Sub KIRYU()
Dim KIR As Range
Dim A As Variant
Dim R As Integer
Dim C As Integer
Dim TARGET As String
   
    TARGET = "発注者品名ｺｰﾄﾞ-備考"
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    A = Array("1013-1400", "1410-1400", "7466-1400")
    
    '品番検索
For Each KIR In Range(Cells(2, C), Cells(Cells(Rows.Count, 2).End(xlUp).Row, C))
    
    If KIR.Value = A(0) Or KIR.Value = A(1) Or KIR.Value = A(2) Then '配列のいずれかと一致した場合の判定
        
        If Workbooks(WWW).Worksheets("DATABASE") _
        .Cells.Find(WHAT:=KIR).Offset(0, 2) = "〇" Then 'DATABASEの中国鋼球欄が〇だったら品番に-中を付ける
        KIR.Value = KIR & "中"
        End If
    ElseIf KIR.Value Like "*-*-*" Or KIR.Offset(0, 3).Value = "ｺﾝﾌﾟ CKD" Or KIR.Offset(0, 4).Value = "KF" Then
    Rows(KIR.Row).Select
    Selection.Font.Strikethrough = True
    
    End If
    
Next
End Sub


Sub SKKSort()
Dim R As Integer
Dim C As Integer
Dim TARGET As String
    
    If UserForm1.ComboBox1.Text = "SKK" Then Columns("A").Insert Shift:=xlToRight
       
    TARGET = "発注者品名ｺｰﾄﾞ-備考"
    
    For AAA = 1 To 2
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    
    ActiveSheet.Range("A1").Sort KEY1:=ActiveSheet.Cells(R, C), ORDER1:=xlAscending, Header:=xlGuess
    
    TARGET = "機種ｺｰﾄﾞ"

    Next


End Sub

Sub SKKEREMA() 'エレマを別ページに移動


Dim R As Integer
Dim C As Integer
Dim TARGET As String

    TARGET = "受渡場所名"
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    
For Each EREMA In Range(Cells(2, C), Cells(Cells(Rows.Count, C).End(xlUp).Row, C))
EREMA.Select '後で消す
    If EREMA.Value = "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
    Rows(EREMA.Row).Select
        
        If EREMA.Offset(1, 0).Value <> "" And EREMA.Offset(1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" _
        And EREMA.End(xlDown).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
        '最終行でなく、1段下がエレマじゃなく、最終行がエレマじゃなかったら最終行へ
        Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA '改ページ
            End If
            ElseIf EREMA.Offset(1, 0).Value <> "" And EREMA.Offset(1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" _
         Then '最終行でなく、1段下がエレマじゃなかったら最終行へ
         Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
            End If
        ElseIf EREMA.Offset(1, 0).Value <> "" And EREMA.End(xlDown).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" _
         Then '最終行でなく、最終行がエレマじゃなかったら最終行へ
         Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
            End If
        ElseIf EREMA.Offset(1, 0).Value = "" And EREMA.Offset(-1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" _
         Then '最終行で、1段上がエレマじゃなかったら改ページ。1注文しかなかったら何もしない
            If EREMA.Offset(-1, 0).Value <> "受渡場所名" Then ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
        ElseIf EREMA.Offset(1, 0).Value = "ｴﾚﾏﾃｯｸ ｳｹｿ" And EREMA.End(xlDown).Value = "ｴﾚﾏﾃｯｸ ｳｹｿ" And EREMA.Offset(-1, 0).Value <> "ｴﾚﾏﾃｯｸ ｳｹｿ" Then
        '1段下がエレマで、1段上がエレマでなく、最終行もエレマの場合改ページ
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
        End If
           
    ElseIf EREMA.Value = "ｻﾝﾜｺｰﾃｯｸｽ" Or EREMA.Value = "ｻｲﾄｰ ｳｹｿｳｺ" Or EREMA.Value = "ﾌｺｸ ｳｹｿｳｺ" Or EREMA.Value = "ｻﾝﾜﾃｯｸ ｳｹｿ" Then
        Rows(EREMA.Row).Select
        Selection.Font.Strikethrough = True

    
        
            
    End If
Next EREMA


'空白削除
For Each EREMA In Range(Cells(2, C), Cells(Cells(Rows.Count, C).End(xlUp).Row, C))
    EREMA.Select
    If EREMA.Value = "" Then
    Rows(EREMA.Row).Delete '2段空白が続いている場合1段ずれるので、対策
    
    End If
Next EREMA

End Sub

