Attribute VB_Name = "Module5"
Option Explicit

Public Sub upData() '�������̍݌ɕ\���쐬
Dim oRow, oCol, i, tRow, tCol, tLastRow, tLastCol As Integer
Dim arr, dataArr As Variant
Dim iRange As Range
Dim lastDay, supplier, current, ye, day, yesNo, cate As String

oRow = bookName.Worksheets("�݌Ɉꗗ").Cells.Find(what:="No", LookAt:=xlWhole).Offset(1, 0).Row
arr = Array("���i�ԍ�", "���i��", "�����݌�", "�����p���b�g��") '�N���E���Y�|�P�b�g
dataArr = Array("���i�ԍ�", "���i��", "�����݌�", "�����p���b�g�݌�") '�N���E���Y�|�P�b�g
tRow = bookName.Worksheets("DATA").Cells.Find(what:="�[�i���t").Offset(1, 0).Row
tCol = bookName.Worksheets("DATA").Cells.Find(what:="�[�i���t").Column
tLastCol = bookName.Worksheets("DATA").Cells.Find(what:="�[�i���t").End(xlToRight).Column
ye = Left(Date, 4)
day = Mid(Date, 6, 2) + 1
cate = Worksheets("�ݒ�").ComboBox2.Text
If day < 10 Then
day = "0" & day
End If
current = bookName.Worksheets("�ݒ�").Range("D3").Text
supplier = bookName.Worksheets("�ݒ�").ComboBox1.Text


If wb(current, supplier, ye, day) = 1 Then

yesNo = MsgBox("�����̏o�ו񍐂�����܂��B�㏑�����܂����H", vbYesNo)

   If yesNo = vbNo Then
       MsgBox ("�o�ו񍐍쐬�𒆒f���܂����B")
       Exit Sub
   End If

End If

lastDay = InputBox("�������i�ŏI�ғ����j�̓��t����͂��Ă��������B" & vbCrLf & "��:2017/5/5  or 5/5")
If lastDay = "False" Then
MsgBox ("�o�ו񍐍쐬�𒆒f���܂����B")
Exit Sub
End If
bookName.Worksheets("�݌Ɉꗗ").Activate

    For i = 0 To 3
    oCol = bookName.Worksheets("�݌Ɉꗗ").Cells.Find(what:=arr(i), LookAt:=xlWhole).Column
    bookName.Worksheets("�݌Ɉꗗ").Range(Cells(oRow, oCol), Cells(oRow, oCol).End(xlDown)).Copy
    bookName.Worksheets("DATA").Cells.Find(what:=dataArr(i)).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Next

bookName.Worksheets("DATA").Activate
tLastRow = bookName.Worksheets("DATA").Cells.Find(what:=dataArr(2)).End(xlDown).Row
bookName.Worksheets("DATA").Range(Cells(tLastRow + 1, tCol), Cells(Rows.Count, tLastCol - 3)).ClearContents

    For Each iRange In Range(Cells(tRow, tCol), Cells(tLastRow, tCol)) '���t�����
    iRange.Value = lastDay
    Next

bookName.Worksheets("�݌Ɉꗗ").Activate

Application.DisplayAlerts = False
bookName.SaveAs Filename:=current & "\" & cate & "�݌Ƀ��X�g" & ye & "." & day & "��.xlsm"
Application.DisplayAlerts = True

Call �s�{�b�g�X�V
Call stockSheet

End Sub

Public Sub stockSheet()
Dim oRow, oCol, eRow, eCol, tCol, tCol2, tLastCol As Integer
oRow = bookName.Worksheets("���P�[�V����").Cells.Find(what:="�i��").Row
oCol = bookName.Worksheets("���P�[�V����").Cells.Find(what:="�i��").Column
eCol = bookName.Worksheets("���P�[�V����").Cells.Find(what:="���݌�").Column
tCol = bookName.Worksheets("���P�[�V����").Cells.Find(what:="�O���p���b�g��").Column
tCol2 = bookName.Worksheets("���P�[�V����").Cells.Find(what:="�o�ɓ�").Column
tLastCol = bookName.Worksheets("���P�[�V����").Cells.Find(what:="�o�Ɍv").Offset(0, -1).Column

With bookName.Worksheets("���P�[�V����")
.Activate
.Cells(oRow, tCol2).Value = Mid(Date, 6, 2) + 1 & "/01"
eRow = .Cells(Rows.Count, oCol).End(xlUp).Row
.Range(Cells(oRow + 1, eCol), Cells(eRow, eCol)).Copy
.Cells(oRow + 1, tCol).PasteSpecial Paste:=xlPasteValues
.Range(Cells(oRow + 1, tCol2), Cells(Rows.Count, tLastCol)).ClearContents
End With

End Sub

Public Function wb(ByVal current, supplier, ye, day As Variant)
Dim buf, defFile As String
Dim A As Integer
A = 0
buf = Dir(current & "\" & "*.xl*")
defFile = "HONDA" & ye & "." & day & "��.xlsm"
Do While buf <> ""
If buf = defFile Then A = A + 1
buf = Dir()
Loop
wb = A
End Function


