VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "�w�����ۑ���A�f�[�^�]�L��w��"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9270
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TEXTFILE As String
Dim TabIndex, ToggleValue As Integer


Private Sub CommandButton2_Click()
Unload UserForm3
End Sub

Private Sub CommandButton3_Click()
Unload UserForm3
End Sub

Private Sub CommandButton4_Click() '�w�����̕ۑ�����w��
On Error GoTo EXS
Dim DIA As FileDialog
 Set DIA = Application.FileDialog(msoFileDialogFolderPicker)
 DIA.Show
 DIA.Title = "�ۑ���̃t�H���_���w�肵�Ă��������B"
 TextBox1.Text = DIA.SelectedItems(1)
Exit Sub
EXS:
End Sub

Private Sub CommandButton5_Click() '�w�����f�[�^�̓]�L����w��
TEXTFILE = Application.GetOpenFilename("*,*.xls*", Title:="�]�L��̃t�@�C�����w�肵�Ă��������B")
    If TEXTFILE = "False" Then
        If TextBox2.Text = "" Then
        TextBox2.Text = TEXTFILE
        End If
    Else
    TextBox2.Text = TEXTFILE
    End If
End Sub

Private Sub CommandButton6_Click()
Worksheets("�ݒ�").Cells(3 + TabIndex, 4).Value = TextBox1.Text
Worksheets("�ݒ�").Cells(3 + TabIndex, 5).Value = TextBox2.Text
    ActiveWorkbook.Save
    CommandButton6.Enabled = False 'Enabled�Ŕ�\��
End Sub

Private Sub TabStrip1_Change()
TabIndex = TabStrip1.Value
TextBox1.Text = Worksheets("�ݒ�").Cells(3 + TabIndex, 4).Value
TextBox2.Text = Worksheets("�ݒ�").Cells(3 + TabIndex, 5).Value
CommandButton6.Enabled = True

End Sub





Private Sub UserForm_Initialize() '�����l�ݒ�
'!���ײ԰���������茸�����肵���ꍇ�̏������l����!
'!�S���Җ����\�L������ω������肳����!
TextBox1.Text = Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TabIndex, 4).Value
TextBox2.Text = Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TabIndex, 5).Value

TabStrip1.Tabs.Item(0).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B3").Value
TabStrip1.Tabs.Item(1).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B4").Value
TabStrip1.Tabs.Item(2).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B5").Value
TabStrip1.Tabs.Item(3).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B6").Value
TabStrip1.Tabs.Item(4).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B7").Value
TabStrip1.Tabs.Item(5).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B8").Value
TabStrip1.Tabs.Item(6).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B9").Value
TabStrip1.Tabs.Item(7).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B10").Value
TabStrip1.Tabs.Item(8).Caption = Workbooks(WWW).Worksheets("�ݒ�").Range("B11").Value
End Sub

