Attribute VB_Name = "Module1"
Option Explicit

'��Ђł�64�r�b�g�Ή��ɂ��� Ptrsafe & longptr
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpCaptionName As String) As LongPtr
Declare PtrSafe Sub SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr)

Public exImCount(3) As Integer

Public Sub callUserForm()
UserForm1.Show
End Sub

Public Sub eICount()
Call ExportAndImportCounter
End Sub


Public Sub ExportAndImportCounter()
Dim wbNum As Integer '�J�E���g���郏�[�N�u�b�N�̐�
Dim wsNum As Integer '�J�E���g���郏�[�N�V�[�g�̐�
Dim i As Integer, l As Integer '���[�v�̕ϐ�
Dim pICount As Integer '�S���Ґ�
Dim monthCount As Integer '�����܂����̍�
Dim startDate As Date, endDate As Date '��������
Dim countBook(3) As String '�����擾�u�b�N�̃p�X
Dim findSheetName As Range '��������V�[�g��
Dim yesNo As String
Dim getTextBook As Workbook '�}�N���̂���V�[�g
Dim hwnd As LongPtr
Dim beforeHwnd As LongPtr
Dim bookPath As Variant, bookName As Variant, pIName() As Variant
Dim pIChange As Range, iRange As Range

'On Error GoTo err

Application.ScreenUpdating = False
Erase exImCount '�z��������l�ɂ���
Set getTextBook = ActiveWorkbook

startDate = InputBox("�����J�n�����w�肵�Ă��������B" & vbCrLf & "�ߋ�1�T�ԕ����ΏۂɂȂ�܂��B", "�����J�n�����͉��", Date - 6)
endDate = startDate + 6

monthCount = Mid(endDate, 6, 2) - Mid(startDate, 6, 2)

For l = 0 To monthCount '�����܂����ꍇ���[�v����

    With getTextBook.Worksheets("�A�o�������擾�V�[�g")
    
        countBook(0) = .Range("D" & 6 + 5 * l).Text '�A���̑䒠
        countBook(1) = .Range("D" & 7 + 5 * l).Text '2�K�֌W�̑䒠
        countBook(2) = .Range("D" & 8 + 5 * l).Text '�A�o�̑䒠
        countBook(3) = .Range("D9").Text '�]�L�Ώۂ̃u�b�N
        
        '�S���҂��Ƃ̌����擾--------
        'Set pIChange = .Cells.Find(what:="�S���Җ�", lookat:=xlWhole).Offset(1)
        'Set pIChange = .Range(Cells(pIChange.Row, pIChange.Column), Cells(pIChange.End(xlDown).Row, pIChange.Column))
        'For Each iRange In pIChange '�S����
            
           ' If iRange.Offset(, 3).Value = "��" Then
                
               'pIName(pICount, 0) = iRange.Value
            
            'End If
            
            'pICount = pICount + 1
            
        'Next
         '----------------------------
          
    End With
    
    For wbNum = 0 To 2 '�u�b�N�̐��������[�v
    
        bookName = Split(countBook(wbNum), "\")
        hwnd = FindWindow("XLMAIN", bookName(UBound(bookName)) & " - Excel")
        
        If hwnd = 0& Then
            Workbooks.Open countBook(wbNum), ReadOnly:=True
        Else
            beforeHwnd = hwnd
            SetForegroundWindow beforeHwnd
        End If
        
        DoEvents
        
        With ActiveWorkbook
            wsNum = .Worksheets.Count
            For i = 1 To wsNum
            'MsgBox Worksheets(i).Name
            
            '�V�[�g�������X�g�ɂȂ��ꍇ-------------------
            Set findSheetName = getTextBook.Worksheets("�A�o�������擾�V�[�g").Columns(1).Find(what:=.Worksheets(i).Name, lookat:=xlWhole)
            If findSheetName Is Nothing Then
                ThisWorkbook.Worksheets("�A�o�������擾�V�[�g").Range("A1").End(xlDown).Offset(1, 0).Value = .Worksheets(i).Name
                yesNo = MsgBox(.Worksheets(i).Name & "�͌�����܂���B" & vbCrLf & "���X�g�ɒǉ����܂����H", vbYesNo)
                If yesNo = vbYes Then
                    ThisWorkbook.Worksheets("�A�o�������擾�V�[�g").Range("A1").End(xlDown).Offset(, 1).Value = "��"
                Else
                    ThisWorkbook.Worksheets("�A�o�������擾�V�[�g").Range("A1").End(xlDown).Offset(, 1).Value = "�~"
                End If
                
            End If
            '-------------------
            
            
                '�擾�ݒ肪"��"�ł���A��\���łȂ����́@�Ifind��after��ݒ�I
                If getTextBook.Worksheets("�A�o�������擾�V�[�g") _
                .Columns(1).Find(what:=.Worksheets(i).Name, lookat:=xlWhole) _
                .Offset(, 1).Value = "��" And .Worksheets(i).Visible = True Then
                    Call getCount(startDate, endDate, Worksheets(i).Name, wbNum)
                End If
                
            Next
            
        End With
        
        ActiveWorkbook.Close (False)
        DoEvents
        
    Next

Next

MsgBox startDate & "�`" & endDate & vbCrLf & "�̗A��������" & _
exImCount(0) & "��" & vbCrLf & "�A��2�K�֌W��" & exImCount(1) & "��" & vbCrLf & "�A�o������" & exImCount(2) & "��" & vbCrLf & "�A�o2�K�֌W��" & exImCount(3) & "���ł��B"

With getTextBook.Worksheets("�A�o�������擾�V�[�g")
    .Range("D1").Value = startDate
    .Range("E1").Value = endDate
    .Range("D2").Value = exImCount(0)
    .Range("E2").Value = exImCount(1)
    .Range("D3").Value = exImCount(2)
    .Range("E3").Value = exImCount(3)
End With

'�]�L����u�b�N���A�N�e�B�u�ɂ���
bookName = Split(countBook(3), "\")
hwnd = FindWindow("XLMAIN", bookName(UBound(bookName)) & " - Excel")

If hwnd = 0& Then
    Workbooks.Open countBook(3)
Else
    beforeHwnd = hwnd
    SetForegroundWindow beforeHwnd
End If

Call tenki(startDate, endDate) '�]�L����

Application.ScreenUpdating = True

Exit Sub

'�G���[����-start-
err:
MsgBox "�G���[�������������߁A�����𒆒f���܂����B" & vbCrLf & err.Number & ":" & err.Description
Application.ScreenUpdating = True
'�G���[����-end-

End Sub

Public Sub getCount(ByVal startDate As Date, ByVal endDate As Date, ByVal sheetName As String, ByVal bookCount As Integer)
Dim numStartRow As Integer, numEndRow As Integer '�����Ώۂ̍ŏI�s
Dim numCol As Integer '"�ԍ�"�̂����
Dim permissionCol As Integer '"����"�̂����
Dim perCount As Integer '�����̃J�E���g
Dim permissionCells As Range '�����͂���Ă���Z���͈�
Dim i As Range

With Worksheets(sheetName)

    .Activate
    numStartRow = .Cells.Find(what:="�ԍ�", lookat:=xlWhole).Row + 1
    numCol = .Cells.Find(what:="�ԍ�", lookat:=xlWhole).Column
    numEndRow = .Cells(numStartRow, numCol).End(xlDown).Row
    permissionCol = .Cells.Find(what:="����", lookat:=xlWhole).Column
    Set permissionCells = .Range(Cells(numStartRow, permissionCol), Cells(numEndRow, permissionCol))
    
        For Each i In permissionCells
            '��������łȂ��A�����J�n���ȏ�A�����I�����ȉ��Ȃ�J�E���g����B
            If i.Value <> "" And i >= startDate And i <= endDate Then
                i.Interior.ColorIndex = 3 '�F�t�����Ċm�F�������Ƃ��B
                perCount = perCount + 1
            End If
            
        Next
        
End With

If ActiveSheet.Name = "�A�o" Then bookCount = 3
exImCount(bookCount) = exImCount(bookCount) + perCount

End Sub

Public Sub tenki(ByVal startDate As Date, ByVal endDate As Date)
Dim selectCol1 As Integer, selectcol2 As Integer
Dim startRow As Integer, endRow1 As Integer, endrow2 As Integer
Dim finalRow As Integer
Dim selectCell1 As Range, selectCell2 As Range, i As Range, dateCell As Range

With ActiveWorkbook.ActiveSheet '�V�[�g���ς�����悤��

    startRow = .Cells.Find(what:="����", lookat:=xlWhole).Row + 1
    selectCol1 = .Cells.Find(what:="����", lookat:=xlWhole).Column
    selectcol2 = .Cells.Find(what:="����", lookat:=xlWhole, after:=.Cells.Find(what:="����")).Column
    endRow1 = .Cells(Rows.Count, selectCol1).End(xlUp).Row
    endrow2 = .Cells(Rows.Count, selectcol2).End(xlUp).Row
    finalRow = 0
    
    Set selectCell1 = .Range(Cells(startRow, selectCol1), Cells(endRow1, selectCol1))
    Set selectCell2 = .Range(Cells(startRow, selectcol2), Cells(endrow2, selectcol2))
    
    '�V�N�x�̃V�[�g�������쐬-----
    Call createSheet(endrow2, selectcol2)
    '-----------------------------
    
    For Each i In selectCell1 '1�ڂ̕\������
    
        If i.Value >= startDate And i <= endDate Then
        
            finalRow = i.Row
            i.Offset(, 3).Value = exImCount(0) + exImCount(1)
            i.Offset(, 4).Value = exImCount(2) + exImCount(3)
            
        End If
        
    Next
    
    If finalRow = 0 Then
    
        For Each i In selectCell2 '2�ڂ̕\������
        i.Select
            If i.Value >= startDate And i <= endDate Then
            
                finalRow = i.Row
                i.Offset(, 3).Value = exImCount(0) + exImCount(1)
                i.Offset(, 4).Value = exImCount(2) + exImCount(3)
                
            End If
            
        Next

    End If
    
    
End With

End Sub

Public Sub createSheet(ByVal endrow2 As Integer, ByVal selectcol2 As Integer) '�V�N�x�̃V�[�g�쐬
Dim beforeDate As Date

With ActiveWorkbook

    If .Worksheets(Worksheets.Count).Cells(endrow2, selectcol2).Offset(0, 4).Value <> "" Then
        
        beforeDate = .Worksheets(Worksheets.Count).Cells(endrow2, selectcol2).Value
        .Worksheets("����").Copy after:=Worksheets(Worksheets.Count)
        .Worksheets(Worksheets.Count).Name = ThisWorkbook.Worksheets("�A�o�������擾�V�[�g").Range("D14").Text
        .Worksheets(Worksheets.Count).Cells.Find(what:="����", lookat:=xlWhole).Offset(1, 0).Value = beforeDate + 1
        
    End If
    
End With

End Sub

Public Sub personInCharge()
Dim inCharge() As String
Dim perRow As Integer, perCol As Integer
Dim perRange As Range, i As Range
Dim personsName As Variant

perRow = ThisWorkbook.Worksheets("�S���Җ���").Rows(1).Find(what:="�c��", lookat:=xlWhole).Row
perRow = ThisWorkbook.Worksheets("�S���Җ���").Rows(1).Find(what:="�c��", lookat:=xlWhole).Column
perRange = ThisWorkbook.Worksheets("�S���Җ���").Range(Cells(perRow, perCol), Cells(Rows.Count, perCol))

For Each i In perRange

    'redim preserve

Next


End Sub


