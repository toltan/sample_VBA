VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�A�o�������擾�Ώۂ̃u�b�N�ݒ���"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
Dim s As String
ChDir ("N:\�Г������ԍ��\")
s = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?", , "�A�o���������擾����u�b�N��I��ł��������B")
If s = "False" Then Exit Sub
Me.TextBox1.Text = s
End Sub

Private Sub CommandButton2_Click()
Dim s As String
ChDir ("N:\�Q�K�֌W\�䒠")
s = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?", , "�A�o���������擾����u�b�N��I��ł��������B")
If s = "False" Then Exit Sub
Me.TextBox2.Text = s
End Sub

Private Sub CommandButton3_Click()
Dim s As String
ChDir ("N:\�Г������ԍ��\")
s = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?", , "�A�o���������擾����u�b�N��I��ł��������B")
If s = "False" Then Exit Sub
Me.TextBox3.Text = s
End Sub

Private Sub CommandButton4_Click()
Dim s As String
ChDir ("N:\�Г������ԍ��\")
s = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?", , "�A�o��������]�L����u�b�N��I��ł��������B")
If s = "False" Then Exit Sub
Me.TextBox4.Text = s
End Sub

Private Sub CommandButton5_Click()
With Worksheets("�A�o�������擾�V�[�g")
    '�O��ݒ肳��Ă����u�b�N�̃p�X��O�񒊏o�Ώۂɓ]�L
    .Range("D11").Value = .Range("D6").Value
    .Range("D12").Value = .Range("D7").Value
    .Range("D13").Value = .Range("D8").Value
    '����ݒ肳�ꂽ�u�b�N�p�X��ݒ�
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
With Worksheets("�A�o�������擾�V�[�g")
    Me.TextBox1.Text = .Range("D6").Text
    Me.TextBox2.Text = .Range("D7").Text
    Me.TextBox3.Text = .Range("D8").Text
    Me.TextBox4.Text = .Range("D9").Text
End With
End Sub
