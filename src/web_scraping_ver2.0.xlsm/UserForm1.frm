VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "日付を選択してください。"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Call fn
End Sub

Private Sub CommandButton2_Click()
Unload UserForm1
Worksheets("scraiping").Activate
End Sub

Private Sub UserForm_Initialize()
Dim DDD As Range
Set DDD = Worksheets(cb1t).Range(Cells(6, 5), Cells(6, Columns.Count).End(xlToLeft))
For Each GGG In DDD
Me.ComboBox1.AddItem (GGG.Text)
Next
End Sub

