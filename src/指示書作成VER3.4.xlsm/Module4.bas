Attribute VB_Name = "Module4"
Sub �]�L(ByVal strNewWorkbookName As String, TABnumber As Integer)
Dim TENKIBOOK As Workbook
Dim newBookname, newSheetname, newSelectSheet As String
Dim WORKSCON, MsgSelect As Integer
Dim TORF As Boolean
    TORF = False
    
    If OpenedBook(TABnumber) = 0 Then '�]�L��t�@�C�����J���Ă��Ȃ�������J��
        If CreateObject("SCRIPTING.FILESYSTEMOBJECT"). _
        FILEEXISTS((Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 5).Value)) = True Then '�t�@�C�������������ꍇ�̏������s��
        Workbooks.Open (Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 5).Value)
        Else
        MsgSelect = MsgBox("�w�肳�ꂽ�t�@�C����������܂���B" & vbCrLf & "�t�@�C���𒼐ڎw�肷�邩�A�I�v�V��������ݒ肵�Ȃ����Ă��������B", vbYesNo, vbExclamation)
            If MsgSelect = vbYes Then
            newSelectSheet = Application.GetOpenFilename("*,*.xlsx", Title:="�]�L��̃t�@�C�����w�肵�Ă��������B")
                If newSelectSheet <> "False" Then
                Workbooks.Open (newSelectSheet)
                Else
                MsgBox ("�L�����Z������܂����B��Ƃ𒆒f���܂��B"), vbExclamation
                Exit Sub
                End If
            Else
            MsgBox ("�L�����Z������܂����B��Ƃ𒆒f���܂��B"), vbExclamation
            Exit Sub
            End If
        End If
    Else
    Workbooks(Dir(Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 5).Value)).Activate
    End If
    
Set TENKIBOOK = ActiveWorkbook
    For Each WS In TENKIBOOK.Worksheets '�����̃V�[�g���Ȃ������ׁA����ꍇ�͒ǉ��B�����ꍇ�͐V�K�V�[�g���쐬����B
        If WS.Name = Mid(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="�[���w���1").Offset(1, 0).Value, 5, 2) _
        & "." & Right(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="�[���w���1").Offset(1, 0).Value, 2) Then
        TORF = True
        newSheetname = WS.Name
        End If
    Next
        If TORF = True Then
        TENKIBOOK.Worksheets(newSheetname).Activate
        Else
        TENKIBOOK.Activate
        WORKSCON = TENKIBOOK.Worksheets.Count
        TENKIBOOK.Worksheets("����").Copy AFTER:=TENKIBOOK.Worksheets(WORKSCON - 1)
        End If
    
    

Call HHHDD(WORKSCON, TORF, strNewWorkbookName, TENKIBOOK)
End Sub


Sub HHHDD(ByVal WORKSCON As Integer, TORF As Boolean, strNewWorkbookName As String, TENKIBOOK As Workbook)

'strNewWorkbookName=�w���� TENKIBOOK=��Ɨp�u�b�N


'!�ǉ������̏ꍇ���l���ɒǉ��ƕt����!
Dim TAR, nextTAR, DAYVVAL As String
Dim R, TARcol, TARrow As Integer
R = 0
TARrow = TENKIBOOK.ActiveSheet.Cells(Rows.Count, TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="���i�ԍ�").Column).End(xlUp).Offset(1, 0).Row
Workbooks(strNewWorkbookName).Activate
For Each RANGVAL In Range("A1:S1") '!TENKIBOOK�ɂ�������TARrow,TARcol�̌��END��OFFSET����̂�����!
    TAR = RANGVAL.Text
    
    If TAR = "�����ҕi������-���l" Or TAR = "�����ҕi������-�[������2" Then '���c�̏ꍇ��҂ɂȂ�
        nextTAR = "���i�ԍ�"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        R = TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).End(xlUp).Offset(1, 0).Row
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "�����ԍ�" Then '!�L�����Ȃ��ꍇ��ɋl�߂Ă��܂̂Œ���!
        nextTAR = "P/O�@No."
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "�[���w������1" Then
        nextTAR = "�o�ɐ���"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="LOT No.").Offset(-1, 0).Value = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="�o�ɐ���").Offset(-1, 0).Value

    
    ElseIf TAR = "��n�ꏊ��" Then
        nextTAR = "�[���ꏊ"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
    
    ElseIf TAR = "�i��(�i���d�l)" Then '�����̂ݕi���L�����������
        If Workbooks(strNewWorkbookName).Worksheets(1).PageSetup.LeftHeader Like "*����*" Then
        nextTAR = "���i��"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        End If
    ElseIf TAR = "�@����" Then '�L�����Ȃ��ꍇ��΂�
        nextTAR = "P/O�@No."
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Offset(1, 2).Column
        If Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp).Value <> RANGVAL.Value Then
        Workbooks(strNewWorkbookName).Worksheets(1).Range(Cells.Find(WHAT:=TAR).Offset(1, 0), Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp).Offset(0, 1)).Copy
        TENKIBOOK.ActiveSheet.Cells(TARrow, TARcol).PasteSpecial Paste:=xlPasteValues
        End If
    ElseIf TAR = "�[���w���1" Then '���i�ԍ��̍s�ԍ����擾���A�[�i���t�ɋL������
      nextTAR = "�[�i���t"
        TARcol = TENKIBOOK.ActiveSheet.Cells.Find(WHAT:=nextTAR).Column
        For Each LTAR In Range(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(1, 0), Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:=TAR).Offset(Rows.Count - 1, 0).End(xlUp))
            If TENKIBOOK.ActiveSheet.Cells(R, TARcol + 1).Value <> "" Then
            TENKIBOOK.ActiveSheet.Cells(R, TARcol).Value = Mid(LTAR.Value, 5, 2) & "/" & Right(LTAR.Value, 2) '���t����
            End If
            R = R + 1
        Next
        '��Ɨp�V�[�g�ɂ���ẮA1�s��̃��C���\��t���ӏ����l����
    End If
Next



DAYVVAL = Mid(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="�[���w���1").Offset(1, 0).Value, 5, 2) _
& "." & Right(Workbooks(strNewWorkbookName).Worksheets(1).Cells.Find(WHAT:="�[���w���1").Offset(1, 0).Value, 2)
If TORF = False Then
TENKIBOOK.Worksheets(WORKSCON).Name = DAYVVAL
End If
TENKIBOOK.Save
End Sub

Sub KLFJDFK()
TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="�o�ɐ���").Offset(-1, 0).Value = _
WorksheetFunction.Sum(TENKIBOOK.ActiveSheet.Range(TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="�o�ɐ���").Offset(1, 0), Cells(Rows.Count, TENKIBOOK.ActiveSheet.Cells.Find(WHAT:="�o�ɐ���").Column).End(xlUp)))
End Sub
