Attribute VB_Name = "Module5"
Option Explicit

Public Function OpenedBook(TABnumber As Integer) '�]�L��u�b�N���J���Ă��邩���ׂ�B
Dim WB As Workbook
Dim A As Long
A = 0

If CreateObject("SCRIPTING.FILESYSTEMOBJECT").FILEEXISTS((Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 5).Value)) = False Then
GoTo NEX
Else
    For Each WB In Workbooks
        If WB.Name = (Dir(Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 5).Value)) Then
        A = A + 1
        End If
    Next
 End If
NEX:
    OpenedBook = A

End Function

Public Function TABno(ByVal SPNAME As String) As Integer '�T�v���C���[�ɂ��w�����̕ۑ����I���B
Dim A, SPno As Integer
For A = 0 To UserForm3.TabStrip1.Tabs.Count - 1
If SPNAME = UserForm3.TabStrip1.Tabs.Item(A).Caption Then
SPno = A
End If
Next
TABno = SPno
End Function

Public Function DIRECT(ByVal SAVRAN As Range, NDAY, NNN As String) As Integer '�����̎w�����������������B
Dim BUF, PS, DEFFILE As String
Dim A As Integer
BUF = Dir(SAVRAN.Value & "\" & "*.xl*")
DEFFILE = (Left(NDAY, 4) & Mid(NDAY, 5, 2) & Mid(NDAY, 7, 2) & Trim(SPNAME) & "�l�[�����w����" & NNN & ".xlsx")
Do While BUF <> ""
If BUF = DEFFILE Then A = A + 1
BUF = Dir()
Loop
DIRECT = A
End Function


Sub HHH()
MsgBox (Month(Now))

End Sub
