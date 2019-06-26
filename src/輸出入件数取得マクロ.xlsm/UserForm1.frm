VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "輸出入件数取得対象のブック設定画面"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
Dim s As String
ChDir ("N:\社内整理番号表")
s = Application.GetOpenFilename("Microsoft Excelブック,*.xls?", , "輸出入件数を取得するブックを選んでください。")
If s = "False" Then Exit Sub
Me.TextBox1.Text = s
End Sub

Private Sub CommandButton2_Click()
Dim s As String
ChDir ("N:\２階関係\台帳")
s = Application.GetOpenFilename("Microsoft Excelブック,*.xls?", , "輸出入件数を取得するブックを選んでください。")
If s = "False" Then Exit Sub
Me.TextBox2.Text = s
End Sub

Private Sub CommandButton3_Click()
Dim s As String
ChDir ("N:\社内整理番号表")
s = Application.GetOpenFilename("Microsoft Excelブック,*.xls?", , "輸出入件数を取得するブックを選んでください。")
If s = "False" Then Exit Sub
Me.TextBox3.Text = s
End Sub

Private Sub CommandButton4_Click()
Dim s As String
ChDir ("N:\社内整理番号表")
s = Application.GetOpenFilename("Microsoft Excelブック,*.xls?", , "輸出入件数を転記するブックを選んでください。")
If s = "False" Then Exit Sub
Me.TextBox4.Text = s
End Sub

Private Sub CommandButton5_Click()
With Worksheets("輸出入件数取得シート")
    '前回設定されていたブックのパスを前回抽出対象に転記
    .Range("D11").Value = .Range("D6").Value
    .Range("D12").Value = .Range("D7").Value
    .Range("D13").Value = .Range("D8").Value
    '今回設定されたブックパスを設定
    .Range("D6").Value = Me.TextBox1.Text
    .Range("D7").Value = Me.TextBox2.Text
    .Range("D8").Value = Me.TextBox3.Text
    .Range("D9").Value = Me.TextBox4.Text
    
End With

End Sub

Private Sub CommandButton6_Click()
Unload UserForm1
End Sub

Private Sub UserForm_Initialize()
With Worksheets("輸出入件数取得シート")
    Me.TextBox1.Text = .Range("D6").Text
    Me.TextBox2.Text = .Range("D7").Text
    Me.TextBox3.Text = .Range("D8").Text
    Me.TextBox4.Text = .Range("D9").Text
End With
End Sub
