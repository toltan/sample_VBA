VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "指示書作成"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
If Me.OptionButton1.Value = True Then Me.OptionButton1.Value = False
If Me.OptionButton2.Value = True Then Me.OptionButton2.Value = False
End Sub

Private Sub UserForm_Initialize()

    With ComboBox1
        .AddItem "正和シール"
        .AddItem "SKK"
        .AddItem "その他(注文少数)"
        .AddItem "黒田複数"
        .AddItem "黒田1枚"
    End With
'''    lngOpenBookNumber = Workbooks.Count
End Sub

Private Sub cmdStart_Click()
Dim CATE As String
    If Me.OptionButton1.Value = True Then
    CATE = "追加-"
    ElseIf Me.OptionButton2.Value = True Then
    CATE = "前倒し-"
    Else
    CATE = ""
    End If

    Call Main(ComboBox1.ListIndex, CATE)

End Sub

Private Sub cmdClose_Click()

    ThisWorkbook.Close
    
End Sub

