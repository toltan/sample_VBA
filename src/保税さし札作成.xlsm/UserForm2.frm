VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "�A�o�����ǂݍ���"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TEXTFILE As String

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

MsgBox "�t�@�C���I��"
ChDir "C:\Users\owner\Desktop"
TEXTFILE = Application.GetOpenFilename("�e�L�X�g �h�L�������g,*.txt")
TextBox1.TEXT = TEXTFILE
End Sub

Private Sub CommandButton1_Click()
Dim OLTTEXT As String
Dim NOTEEXCEL As String
OLTTEXT = TextBox1.TEXT
NOTEEXCEL = ("TEXT;" & OLTTEXT)

If TextBox1.TEXT = "" Then
MsgBox "���͂Ȃ�"
Exit Sub
Else
Worksheets("�N�G��").Cells.Delete
 With Worksheets("�N�G��").QueryTables.Add(Connection:=NOTEEXCEL, Destination:=Worksheets("�N�G��").Range("$B$1"))
        .Name = Dir(OLTTEXT)
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 1, 1)
        .TextFileFixedColumnWidths = Array(80, 10)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveWindow.SmallScroll Down:=-30
    
Unload UserForm2
UserForm1.Show (vbModeless)


End If
End Sub
