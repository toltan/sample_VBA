Attribute VB_Name = "Module3"
Option Explicit
Public bookName As Workbook

Public Sub WO(ByVal A As String)
Set bookName = Workbooks(A)
End Sub

Public Sub HONDA() '�z���_���X�g�ɓ��ɂ�ǉ�������o�ɂ���������B�����B

Dim ws As Worksheet
Dim i As Range
Dim iRange As Object
Dim VAR As Variant

Set ws = Worksheets("DATA")

    If ws.AutoFilterMode = False Then
    Rows(ws.Range("B3").Row).AutoFilter
    End If
    
ws.Range("B3").AutoFilter FIELD:=2, Criteria1:="2017/4/10" '���t�͕ϐ��ݒ�

Set iRange = ws.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)

For Each i In iRange

Next
End Sub

Public Sub newPartsNumber()
Dim catNo, oRow, oCol, tRow, tCol, tColZ, i, nRow, bool, inser, iRow, D As Integer
Dim cat, regi As Variant
Dim yesNo3 As String

cat = Array("���i�ԍ�", "���i��", "�O���c", "�O���p���b�g��")
catNo = 0
tRow = Worksheets("�V�i�Ԓǉ�").Cells.Find(what:="���i�ԍ�").Offset(1, 0).Row
tCol = Worksheets("�V�i�Ԓǉ�").Cells.Find(what:="���i�ԍ�").Offset(1, 0).Column
tColZ = Worksheets("�V�i�Ԓǉ�").Cells.Find(what:="�O���c").Offset(1, 0).Column
bool = 0
inser = 0
D = 0

Worksheets("�V�i�Ԓǉ�").Activate
Worksheets("�V�i�Ԓǉ�").Range(Cells(tRow, tCol), Cells(Rows.Count, Columns.Count)).ClearContents



For i = 0 To 3 '�݌ɕ\�̕��i�ԍ��A���i���A�O���c�̍s��V�i�Ԓǉ��V�[�g�ɓ]�L
    Worksheets("�݌Ɉꗗ").Activate
    oRow = Worksheets("�݌Ɉꗗ").Cells.Find(what:=cat(catNo)).Offset(1, 0).Row
    oCol = Worksheets("�݌Ɉꗗ").Cells.Find(what:=cat(catNo)).Offset(1, 0).Column
    Worksheets("�݌Ɉꗗ").Range(Cells(oRow, oCol), Cells(oRow, oCol).End(xlDown)).Copy
    Worksheets("�V�i�Ԓǉ�").Cells.Find(what:=cat(catNo)).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    catNo = catNo + 1
    
Next

Worksheets("�V�i�Ԓǉ�").Activate
Worksheets("�V�i�Ԓǉ�").Range(Cells(tRow, tCol).End(xlDown).Offset(1, 2), Cells(Rows.Count, tColZ).Offset(0, tColZ)).ClearContents
nRow = Worksheets("�V�i�Ԓǉ�").Cells(Rows.Count, tCol).End(xlUp).Offset(1, 0).Row '�V�i�Ԓǉ��V�[�g�̍ŏI�s���擾
  
With Worksheets("�V�i�Ԓǉ�") '�V�i�ԓ���

Do While bool >= 0
     
    regi = registration() 'Function registration���i�ԁA�i�����擾

    If regi(0) = "cancel" Then '�i�ԁA�i���ǂ��炩�������͂̏ꍇ�̓L�����Z��
    MsgBox ("�L�����Z�����܂����B")
    Exit Sub
    ElseIf regi(0) = "null" Then
    MsgBox ("�i�Ԃ������ׁ͂̈A�L�����Z�����܂����B")
    Exit Sub
    End If

    .Cells(nRow + bool, tCol).Value = regi(0) '�V�i�Ԃ�ǉ�
    .Cells(nRow + bool, Cells.Find(what:="���i��").Column).Value = regi(1)
    .Cells(nRow + bool, Cells.Find(what:="�O���c").Column).Value = "0"
    .Cells(nRow + bool, Cells.Find(what:="�O���p���b�g��").Column).Value = "0"
    
    yesNo3 = MsgBox("�����ĐV�i�ԓo�^���s���܂����H", vbYesNo + vbQuestion) '�Q�ȏ�̐V�i�Ԃ�����A�A�����ēo�^����Ȃ�YES��LOOP
    
    If yesNo3 = vbYes Then
    bool = bool + 1
    inser = inser + 1
    Else
    bool = -1
    End If

Loop
    
    If .AutoFilterMode = False Then '�I�[�g�t�B���^�[���|�����ĂȂ���Ί|����
        .Rows(4).AutoFilter
    End If
    .Range(Cells(tRow - 1, tCol), Cells(Rows.Count, Cells(tRow - 1, tCol).End(xlToRight).Column + 1)).Sort KEY1:=Cells(tRow + 1, tCol), ORDER1:=xlAscending, Header:=xlYes
    '���ёւ�����

End With

iRow = Worksheets("DATA").Cells.Find(what:="�����݌�").End(xlDown).Offset(1, 0).Row

Worksheets("DATA").Range(iRow & ":" & iRow + inser).Insert '�V�[�g�i�݌Ɉꗗ�j�A�iDATA�j�ɐV�i�Ԃ�ǉ�
Worksheets("�V�i�Ԓǉ�").Range(Cells(tRow, tCol), Cells(Rows.Count, tColZ + 2).End(xlUp)).Copy
Worksheets("DATA").Cells.Find(what:="�[�i���t").Offset(1, 1).PasteSpecial Paste:=xlPasteValues

    Do Until D = 1 + inser '�V�[�g�iDATA�j�̑}�������s�Ɍ����̓��ɂ���}�������s�̕���������
    Worksheets("DATA").Cells.Find(what:="�[�i���t").End(xlDown).Offset(1, 0).Value = _
    Worksheets("DATA").Cells.Find(what:="�[�i���t").End(xlDown).Value
    D = D + 1
    Loop

Call stockPlus(D)
Call stockListPlus(D)
Worksheets("�V�i�Ԓǉ�").Activate
Worksheets("�V�i�Ԓǉ�").Range(Cells(tRow, tCol), Cells(Rows.Count, tCol).End(xlUp)).Copy
Worksheets("�݌Ɉꗗ").Cells.Find(what:="���i�ԍ�").Offset(1, 0).PasteSpecial Paste:=xlPasteValues
Worksheets("�V�i�Ԓǉ�").Range(Cells(tRow, Cells.Find(what:="���i��").Column), Cells(Rows.Count, Cells.Find(what:="���i��").Column).End(xlUp)).Copy
Worksheets("�݌Ɉꗗ").Cells.Find(what:="���i��").Offset(1, 0).PasteSpecial Paste:=xlPasteValues

Call �s�{�b�g�X�V
End Sub

Public Function registration() '�V�i�ԓo�^
Dim A As Boolean
Dim newPart, newPartName, yesNo, yesNo2, oKOnly As String
Dim newP(1) As Variant
A = False

    Do While A = False
    
        newPart = InputBox("�V�i�Ԃ���͂��Ă��������B")
        
        If newPart = "" Then
            newPart = "null"
            Exit Do
        End If
        
        newPartName = InputBox("�V�i�Ԃ̕i������͂��Ă��������B")
        
        If newPartName = "" Then
            newPart = "null"
            Exit Do
        End If
        
        yesNo = MsgBox("�i��:" & newPart & " " & "�i��:" & newPartName & vbCrLf & "�ȏ�̓��e�œo�^���܂��B�X�����ł����H", vbYesNo)
        
            If yesNo = vbNo Then
            A = False
            yesNo2 = MsgBox("�V�K�i�ԓo�^���L�����Z�����܂����H", vbYesNo)
            
                If yesNo2 = vbYes Then
                newPart = "cancel"
                Exit Do
                End If
                
            Else
            A = True
            End If
            
    Loop
    
newP(0) = newPart
newP(1) = newPartName
registration = newP

End Function

Public Sub stockPlus(ByVal D As Integer) '�݌ɁA�[�i���s�𑝂₷

Dim lastRow, newLastCol, newLastRow As Integer
Dim i As Integer
Dim newLastRowVal As String
Dim pivotRange As Range
Dim splitPivot As Variant

Worksheets("�݌ɁA�[�i").Activate

For i = 1 To D

    lastRow = Worksheets("�݌ɁA�[�i").Cells.Find(what:="����").End(xlDown).Row 'A�񂪉B��Ă�
    newLastCol = Cells(Rows.Count, Cells.Find(what:="���i�ԍ�").Column).End(xlUp).Column
    newLastRow = Cells(Rows.Count, Cells.Find(what:="���i�ԍ�").Column).End(xlUp).Row
    newLastRowVal = Cells(Rows.Count, Cells.Find(what:="���i�ԍ�").Column).End(xlUp).Value
    Set pivotRange = Worksheets("�s�{�b�g").Cells.Find(what:="*****")
    
    With Worksheets("�݌ɁA�[�i")
    
        .Range(lastRow - 5 & ":" & lastRow).Copy
        .Range(lastRow + 1 & ":" & lastRow + 1).PasteSpecial
        
    End With
    
    splitPivot = Split(pivotRange.Address, "$") 'Address�ŃZ���ԍ����擾��"$"�ŋ�؂�z���
    newLastRowVal = "=�s�{�b�g!" & splitPivot(1) & newLastRow + 1 '��������
    Worksheets("�݌ɁA�[�i").Cells(newLastRow + 6, newLastCol).Value = newLastRowVal
    
Next

End Sub

Public Sub stockListPlus(ByVal D As Integer) '�݌Ɉꗗ���s�𑝂₷ D�ŉ�
Dim noCol, noRow, noEndRow, noEndCol, ist As Integer
Dim i, stockRange, stockRange2 As Range
Dim stockNo, stockNo2, cellNo As Variant
Dim nouba As String
With Worksheets("�݌Ɉꗗ")
    
    .Activate
    noCol = .Cells.Find(what:="No", LookAt:=xlWhole).Column
    noRow = .Cells.Find(what:="No", LookAt:=xlWhole).Row
    Set stockRange = Worksheets("�݌ɁA�[�i").Cells.Find(what:="�O���c")
    stockNo = Split(stockRange.Address, "$")
    Set stockRange2 = Worksheets("�݌ɁA�[�i").Cells.Find(what:="���v")
    stockNo2 = Split(stockRange2.Address, "$")
    
End With
    
For ist = 1 To D

    noEndRow = Worksheets("�݌Ɉꗗ").Cells(noRow, noCol).End(xlDown).Offset(1, 0).Row '�݌Ɉꗗ�̍ŏI�s���擾
    noEndCol = Worksheets("�݌Ɉꗗ").Cells(noRow, Columns.Count).End(xlToLeft).Column '�݌Ɉꗗ�̍ŏI����擾
    Worksheets("�݌Ɉꗗ").Rows(noEndRow).Insert

    
    For Each i In Range(Cells(noEndRow, noCol), Cells(noEndRow, noEndCol))
        If i.Offset(-1, 0).End(xlUp).Value = "No" Then
        i.Value = i.Offset(-1, 0).Value + 1
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�O���c" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo(1) & cellNo(1) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�O���p���b�g��" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo(1) & cellNo(1) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "���ɗ݌v" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "���Ƀp���b�g�݌v" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�o�ɗ݌v" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�o�Ƀp���b�g�݌v" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�����݌�" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�����p���b�g��" Then
        cellNo = Split(i.Offset(-1, 0).Formula, stockNo2(1))
        i.Value = "='�݌ɁA�[�i'!" & stockNo2(1) & cellNo(UBound(cellNo)) + 6
        ElseIf i.Offset(-1, 0).End(xlUp).Value = "�[��" Then
        i.Value = "-"
        End If
        
    Next
    
Next



End Sub





