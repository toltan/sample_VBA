VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "各項目入力フォーム"
   ClientHeight    =   12870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10365
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
ComboBox1.TEXT = ""
ComboBox2.TEXT = ""
TextBox1.TEXT = ""
TextBox2.TEXT = ""
TextBox3.TEXT = ""
TextBox4.TEXT = ""
TextBox5.TEXT = ""
ComboBox3.TEXT = "C"
ComboBox3.TEXT = ""
ComboBox4.TEXT = ""
TextBox6.TEXT = ""
TextBox7.TEXT = ""
ComboBox5.TEXT = ""
TextBox8.TEXT = ""
End Sub

Private Sub CommandButton4_Click()
Unload UserForm1
End Sub

Private Sub UserForm_Initialize()
Dim C As String, D As String
Dim i As Variant
Dim sup, packing As Range
A = Worksheets("チェックシート").Range("H17").Value
Set sup = Range(Worksheets("設定").Cells.Find(WHAT:="サプライヤー名").Offset(1, 0), Worksheets("設定").Cells(Rows.Count, Worksheets("設定").Cells.Find(WHAT:="サプライヤー名").Column).End(xlUp))
Set packing = Range(Worksheets("設定").Cells.Find(WHAT:="荷姿単位").Offset(1, 0), Worksheets("設定").Cells(Rows.Count, Worksheets("設定").Cells.Find(WHAT:="荷姿単位").Column).End(xlUp))

With UserForm1.ComboBox1
    For Each i In sup
        .AddItem i.Value
    Next
End With

With UserForm1.ComboBox2
    For Each i In sup
        .AddItem i.Value
    Next
End With

With UserForm1.ComboBox3
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
End With

With UserForm1.ComboBox4
    For Each i In sup
        .AddItem i.Value
    Next
End With

With UserForm1.ComboBox5
    For Each i In packing
        .AddItem i.Value
    Next
End With

With UserForm1.ComboBox6
   For Each i In packing
        .AddItem i.Value
    Next
End With

With UserForm1.ComboBox7
    For Each i In packing
        .AddItem i.Value
    Next
End With



ComboBox1.TEXT = Worksheets("さし札").Range("B4").Value
ComboBox2.TEXT = Worksheets("さし札").Range("B7").Value
TextBox1.TEXT = Worksheets("さし札").Range("A11").Value
TextBox2.TEXT = Worksheets("さし札").Range("C11").Value
TextBox3.TEXT = Worksheets("さし札").Range("F11").Value
TextBox4.TEXT = Worksheets("さし札").Range("A14").Value
TextBox5.TEXT = Worksheets("さし札").Range("C14").Value
ComboBox3.TEXT = Worksheets("さし札").Range("A18").Value
ComboBox4.TEXT = Worksheets("さし札").Range("C18").Value
TextBox6.TEXT = Worksheets("さし札").Range("A21").Value
TextBox7.TEXT = Worksheets("さし札").Range("B23").Value
ComboBox5.TEXT = Worksheets("さし札").Range("D23").Value
TextBox8.TEXT = Worksheets("さし札").Range("F23").Value
TextBox10.TEXT = Worksheets("さし札").Range("B70").Value
TextBox11.TEXT = Worksheets("チェックシート").Range("E23").Value
TextBox12.TEXT = Worksheets("チェックシート").Range("E25").Value
TextBox13.TEXT = Worksheets("チェックシート").Range("G23").Value
TextBox14.TEXT = Worksheets("チェックシート").Range("G25").Value
TextBox15.TEXT = Worksheets("チェックシート").Range("I23").Value
TextBox16.TEXT = Worksheets("チェックシート").Range("I25").Value
TextBox9.TEXT = Left(A, 4) & Mid(A, 6, 2) & Mid(A, 9, 2)
TextBox10.TEXT = Mid(A, 12, 4) & Mid(A, 17, 2) & Mid(A, 20, 2)
ToggleButton1.Value = False

End Sub


Private Sub ComboBox3_Change()
If ComboBox3.Value = "B" Then
ToggleButton1.Value = True
Else
ToggleButton1.Value = False
End If
End Sub

Private Sub ToggleButton1_Click()
If ToggleButton1.Value = True Then
Label4.Caption = ("コンテナ番号")
ComboBox3.TEXT = ("B")
Else
Label4.Caption = ("貨物管理番号")
ComboBox3.TEXT = ("C")
End If

End Sub


Private Sub CommandButton1_Click()
 Dim A As String, B As String
 A = TextBox9.TEXT
 B = TextBox10.TEXT
 Dim PPP As String
 PPP = Application.ActivePrinter
 Application.ActivePrinter = "KONICA MINOLTA 423SeriesPCL on Ne02:"
 With Worksheets("さし札")
.Range("B4").Value = ComboBox1.TEXT
.Range("B7").Value = ComboBox2.TEXT
.Range("C11").Value = TextBox2.TEXT
.Range("F11").Value = TextBox3.TEXT
.Range("A14").Value = TextBox4.TEXT
.Range("C14").Value = TextBox5.TEXT
.Range("A18").Value = ComboBox3.TEXT
.Range("C18").Value = ComboBox4.TEXT
.Range("A21").Value = TextBox6.TEXT
.Range("B23").Value = TextBox7.TEXT
.Range("D23").Value = ComboBox5.TEXT
.Range("F23").Value = TextBox8.TEXT
.Range("C10").Value = Label4.Caption
.Range("F11").Value = TextBox3.TEXT
.Range("A4:H24").Copy
.Range("J4:Q24").PasteSpecial
End With
ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False

If CheckBox1.Value = True Then
 With Worksheets("さし札2")
.Range("B4").Value = ComboBox1.TEXT
.Range("B7").Value = ComboBox2.TEXT
.Range("C11").Value = TextBox1.TEXT
.Range("F11").Value = TextBox21.TEXT
.Range("A14").Value = TextBox4.TEXT
.Range("C14").Value = TextBox19.TEXT
.Range("A18").Value = ComboBox3.TEXT
.Range("C18").Value = ComboBox4.TEXT
.Range("A21").Value = TextBox20.TEXT
.Range("B23").Value = TextBox17.TEXT
.Range("D23").Value = ComboBox6.TEXT
.Range("F23").Value = TextBox22.TEXT
.Range("C10").Value = Label4.Caption
.Range("F11").Value = TextBox21.TEXT
.Range("A4:H24").Copy
.Range("J4:Q24").PasteSpecial
End With
Worksheets("さし札2").PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False
End If

If CheckBox1.Value = False Then

Worksheets("チェックシート").Select
Range("I1").Value = "区分:" & ComboBox4.TEXT
Range("H13").Value = TextBox4.TEXT
Range("H17").Value = Left(A, 4) & "/" & Mid(A, 5, 2) & "/" & Right(A, 2) _
 & "～" & Left(B, 4) & "/" & Mid(B, 5, 2) & "/" & Right(B, 2)
Range("E23").Value = TextBox11.TEXT
Range("E25").Value = TextBox12.TEXT
Range("G23").Value = TextBox13.TEXT
Range("G25").Value = TextBox14.TEXT
Range("I23").Value = TextBox15.TEXT
Range("I25").Value = TextBox16.TEXT
Range("D45").Value = TextBox2.TEXT
Range("I37").Value = "計" & "_________" & "(" & TextBox7.TEXT & ")" & " " & ComboBox5.TEXT
ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False
Else
Worksheets("チェックシート").Select
Range("I1").Value = "区分:" & ComboBox4.TEXT
Range("H13").Value = TextBox4.TEXT
Range("H17").Value = Left(A, 4) & "/" & Mid(A, 5, 2) & "/" & Right(A, 2) _
 & "～" & Left(B, 4) & "/" & Mid(B, 5, 2) & "/" & Right(B, 2)
Range("E23").Value = TextBox11.TEXT
Range("E25").Value = TextBox12.TEXT
Range("D45").Value = TextBox2.TEXT
Range("I37").Value = "計" & "_________" & "(" & TextBox7.TEXT & ")" & " " & ComboBox5.TEXT
ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False
        
Worksheets("チェックシート2").Select
Range("I1").Value = "区分:" & ComboBox4.TEXT
Range("H13").Value = TextBox4.TEXT
Range("H17").Value = Left(A, 4) & "/" & Mid(A, 5, 2) & "/" & Right(A, 2) _
 & "～" & Left(B, 4) & "/" & Mid(B, 5, 2) & "/" & Right(B, 2)
Range("E23").Value = TextBox13.TEXT
Range("E25").Value = TextBox14.TEXT
Range("D45").Value = TextBox1.TEXT
Range("I37").Value = "計" & "_________" & "(" & TextBox17.TEXT & ")" & " " & ComboBox6.TEXT
ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False
        
End If

Worksheets("さし札").Select
Range("A1").Select
Application.ActivePrinter = PPP

Unload UserForm1

End Sub

Private Sub 読み取り_Click()
ComboBox1.TEXT = Worksheets("クエリ").Range("B46").Value
ComboBox2.TEXT = Worksheets("クエリ").Range("B46").Value

If CheckBox1.Value = True Then
TextBox1.TEXT = Worksheets("クエリ").Range("B66").Value
TextBox19.TEXT = Worksheets("クエリ").Range("B70").Value
TextBox20.TEXT = Worksheets("クエリ").Range("B69").Value
TextBox17.TEXT = Worksheets("クエリ").Range("B88").Value
ComboBox6.TEXT = Worksheets("クエリ").Range("B89").Value
TextBox21.TEXT = Worksheets("クエリ").Range("B76").Value & " " & Worksheets("クエリ").Range("B77").Value
End If
TextBox4.TEXT = Worksheets("クエリ").Range("B6").Value
TextBox2.TEXT = Worksheets("クエリ").Range("B32").Value
TextBox6.TEXT = Worksheets("クエリ").Range("B35").Value
TextBox5.TEXT = Worksheets("クエリ").Range("B36").Value
TextBox7.TEXT = Worksheets("クエリ").Range("B54").Value
ComboBox5.TEXT = Worksheets("クエリ").Range("B55").Value
If CheckBox1.Value = True Then
TextBox9.TEXT = Worksheets("クエリ").Range("B205").Value
Else
TextBox9.TEXT = Worksheets("クエリ").Range("B69").Value
End If
If CheckBox1.Value = True Then
TextBox10.TEXT = Worksheets("クエリ").Range("B206").Value
Else
TextBox10.TEXT = Worksheets("クエリ").Range("B70").Value
End If
If CheckBox1.Value = True Then
TextBox11.TEXT = Worksheets("クエリ").Range("B211").Value
TextBox12.TEXT = Worksheets("クエリ").Range("B213").Value
TextBox13.TEXT = Worksheets("クエリ").Range("B220").Value
TextBox14.TEXT = Worksheets("クエリ").Range("B222").Value
TextBox15.TEXT = Worksheets("クエリ").Range("B229").Value
TextBox16.TEXT = Worksheets("クエリ").Range("B231").Value
Else
TextBox11.TEXT = Worksheets("クエリ").Range("B75").Value
TextBox12.TEXT = Worksheets("クエリ").Range("B77").Value
TextBox13.TEXT = Worksheets("クエリ").Range("B84").Value
TextBox14.TEXT = Worksheets("クエリ").Range("B86").Value
TextBox15.TEXT = Worksheets("クエリ").Range("B93").Value
TextBox16.TEXT = Worksheets("クエリ").Range("B95").Value
End If
TextBox3.TEXT = Worksheets("クエリ").Range("B42").Value & " " & Worksheets("クエリ").Range("B43").Value
End Sub
