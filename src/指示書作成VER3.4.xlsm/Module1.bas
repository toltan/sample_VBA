Attribute VB_Name = "Module1"
Option Explicit '2015/08/23�쐬
'�V�K���[�N�u�b�N��
Dim strNewWorkbookName As String  '�]�L�� �w����
Dim strCsvbookName, NNN As String '�]�L�� CSV
Dim lngRowCounter As Long
Dim lngCellCounter As Long
Dim SAVRAN As Range
Public KA, KC, TABnumber As Integer
Public lngOpenBookNumber As Long
Public SPNAME, WWW, addNo As String
Public SW As Byte
Public WB1, WB2, WB3 As Workbook

Function f_NewWorkbook() As String

    '�V����Book�I�[�v��
    Workbooks.Add
    f_NewWorkbook = ActiveWorkbook.Name
    
End Function

Public Sub Main(ByVal valListIndex As Variant, CATE As String)
SW = 0
NNN = ""
addNo = "1"

    If valListIndex = -1 Then
        MsgBox "���X�g����I�����Ă�������"
        Exit Sub
    End If

    '�t�@�C���I�[�v��
    If valListIndex = 3 Then
    Call ���c����
    Else
    Call OpenFile
    End If
    'CSV�t�@�C�����̃Z�b�g
    strNewWorkbookName = f_NewWorkbook
    
    lngOpenBookNumber = Workbooks.Count
    
    '���a�V�[��
    If valListIndex = 0 Or valListIndex = 2 Then
        Call �����ԍ�
        Call ���i��
        Call �����ҕi��
        Call �[���w����
        Call �[���w���
'''20141111�ǉ�
        Call �[������
        
        Call ��n�ꏊ
        Call �s����
        Call �݌v�ύX�ԍ�
        Call �@��R�[�h
        Call �������喼
        
        Call ���ёւ����a
        '20160111
        Call �폜(valListIndex)
    'SKK
   ElseIf valListIndex = 1 Then
        Call �����ԍ�
        Call ���i��
        Call �����ҕi��
        Call �[���w����
        Call �[���w���
'''20141111�ǉ�
        Call �[������
        
        Call ��n�ꏊ
        Call �s����
        Call �݌v�ύX�ԍ�
        Call �@��R�[�h
        Call �������喼
        
       
        '20160111
        Call �폜(valListIndex)
        '20161104���c�ǉ�
    ElseIf valListIndex = 3 Or valListIndex = 4 Then
        Call �����ԍ�
        Call ���i��
        Call �����ҕi��
        Call �[���w����
        Call �[���w���
        Call �[������
        Call KURODAUKEWATASHI
        
    
    
    
        
    '�I������
    Else
        MsgBox "���X�g����I�����Ă�������"
        Exit Sub
    End If
    
    '�ǉ�����
    Dim AAAF As Integer
    
    If UserForm1.ComboBox1.Text = "���a�V�[��" Or UserForm1.ComboBox1.Text = "SKK" Then
    AAAF = 28
    ActiveSheet.PageSetup.PaperSize = xlPaperB4
    ActiveWindow.Zoom = 70
   
    Call KDODNA
    
    ElseIf UserForm1.ComboBox1.Text = "SKK" Then
    AAAF = 25
    ActiveSheet.PageSetup.PaperSize = xlPaperB4
    ActiveWindow.Zoom = 70
    Else
    AAAF = 35
    
    End If
    
    Cells.Select
    Selection.RowHeight = AAAF
        
        '���ײ԰���擾
        Dim FIFI As Range
        Dim HINBANUP As String
        Dim FCOL As Integer
        Dim FROW As Integer
        Dim VBS As Integer
        Dim ACS As Worksheet
        
        
        Set ACS = ActiveWorkbook.ActiveSheet
        If valListIndex = 3 Or valListIndex = 4 Then  '���c�̏ꍇ�i�ԂP��̍��ږ����Ⴄ��
        HINBANUP = "�����ҕi������-�[������2"
        Else
        HINBANUP = "�����ҕi������-���l"
        End If
        Set FIFI = Workbooks(WWW).Worksheets("DATABASE").Cells.Find _
        (WHAT:=ACS.Cells.Find(HINBANUP).End(xlDown).Value, LookIn:=xlValues, LOOKAT:=xlWhole)
            If FIFI Is Nothing Then
            VBS = MsgBox(ACS.Cells.Find(HINBANUP).End(xlDown).Value & "��������܂���ł����B�i�Ԃ�ǉ����Ă���Ďn�����邩�A" & vbCrLf & "���ײ԰�����L�����Ă��������B", vbOKCancel)
                 
                If VBS = vbOK Then
                UserForm2.Show (vbModal)
                If SW = 3 Then '�t�H�[���̃L�����Z���{�^���������ꂽ��I������B
                Workbooks(strCsvbookName).Close
                GoTo ErrorHandler
                End If
                GoTo SNA
                ElseIf VBS = vbCancel Then
                Workbooks(strCsvbookName).Close
                GoTo ErrorHandler
                End If
            End If
        
        FCOL = FIFI.Column
        FROW = FIFI.Row
        SPNAME = Workbooks(WWW).Worksheets("DATABASE").Cells(FROW, FCOL).End(xlUp).Value
        TABnumber = TABno(SPNAME)
SNA:
        TABnumber = TABno(SPNAME)

        '�w�b�_�[�A�t�b�^�[�A�s�����𒲐�
        With ActiveSheet.PageSetup
        Dim NDAY As String
        NDAY = ACS.Cells.Find("�[���w���1").Offset(1, 0).Value
        .Orientation = xlLandscape
        If valListIndex <> 3 And valListIndex <> 4 Then
        .LeftHeader = "&13 " & "&B" & Mid(NDAY, 5, 2) & "/" & Mid(NDAY, 7, 2) & "  " & SPNAME
        Else
        .LeftHeader = "&13 " & "&B" & Mid(NDAY, 5, 2) & "/" & Mid(NDAY, 7, 2) & "  " & SPNAME & vbCr & KA & "/" & KC & "��"
        End If
        .RightHeader = "&B" & "&P" & "/" & "&N"
        End With
        Columns("A:A").ColumnWidth = 6
        Columns(ACS.Cells.Find("�����ԍ�").Column).AutoFit '�i���A�i�ԁA��n���ꏊ���̗�𒲐�����悤�Ƀ}�N������2016/09/05
        Columns(ACS.Cells.Find("�[���w���1").Column).ColumnWidth = 8.5 '����
        Columns(ACS.Cells.Find("�i��(�i���d�l)").Column).ColumnWidth = 6 '�i��
        Columns(ACS.Cells.Find("��n�ꏊ��").Column).ColumnWidth = 6
        Columns(ACS.Cells.Find("�[���w���1").Column).ColumnWidth = 11.5
        Columns(ACS.Cells.Find("�[���w������1").Column).ColumnWidth = 7.5
        Columns(ACS.Cells.Find(HINBANUP).Column).ColumnWidth = 15
        Columns("G:K").ColumnWidth = 5.5 '��n�ꏊ�ȍ~
        
        Select Case SPNAME
        Case "���a�@��H��"
            Call SKKSort
            Call KEISEN(SPNAME)
        
            Columns(ACS.Cells.Find("�����ҕi������-���l").Column).ColumnWidth = 20
            Columns(ACS.Cells.Find("��n�ꏊ��").Column).ColumnWidth = 12
            Columns("A:A").ColumnWidth = 3

            Call SKKEREMA
        Case "���a�V�[���̔�"
        If UserForm1.ComboBox1.Text = "���a�V�[��" Then
            Columns(ACS.Cells.Find("�[������1").Column).Delete
        End If
            Call KEISEN(SPNAME)
        
        Case "�O�a�@�ː��H��"
            Call KIRYU
            Call KEISEN(SPNAME)
        Case "�������쏊"
            Call KEISEN(SPNAME)
        Case "�O�a�@�O���H��"
            Call KEISEN(SPNAME)
            Call MAEBASI
        Case "�����m�H��"
            Call KEISEN(SPNAME)
        Case "�G���}�e�b�N"
            Call KEISEN(SPNAME)
        Case "���c���쏊"
            Cells.RowHeight = 20
            Call KURODAKEISEN
            
        End Select
        
    '�Q�ƌ�CSV�t�@�C���̃N���[�Y
    Workbooks(strCsvbookName).Close SaveCHANGES:=False
        
        If Dir(Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 4).Value, vbDirectory) = "" Then '�w�肳�ꂽ�p�X�������ꍇ�̓f�X�N�g�b�v�ɕۑ��B
        MsgBox ("�w�肳�ꂽ�t�H���_��������܂���ł����B" & vbCrLf & "�f�X�N�g�b�v�ɕۑ����܂��B"), vbInformation
        Set SAVRAN = Workbooks(WWW).Worksheets("�ݒ�").Range("D100")
        Else
        Set SAVRAN = Workbooks(WWW).Worksheets("�ݒ�").Cells(3 + TABnumber, 4)
        End If
        
        
        Dim SheetNo As String
        SheetNo = (Left(NDAY, 4) & Mid(NDAY, 5, 2) & Mid(NDAY, 7, 2) & Trim(SPNAME) & "�l�[�����w����" & NNN & ".xlsx")
            If CATE <> "" Then
            
            NNN = CATE & addNo '���ڂ�ǉ�
                
                Do While DIRECT(SAVRAN, NDAY, NNN) = 1 '�����̃V�[�g������������
                addNo = addNo + 1 '�����̃V�[�g���������疼�O��ς��čČ���
                NNN = CATE & addNo
                SheetNo = (Left(NDAY, 4) & "." & Mid(NDAY, 5, 2) & "." & Mid(NDAY, 7, 2) & Trim(SPNAME) & "�l�[�����w����" & NNN & ".xlsx")
                Loop
            SheetNo = (Left(NDAY, 4) & Mid(NDAY, 5, 2) & Mid(NDAY, 7, 2) & Trim(SPNAME) & "�l�[�����w����" & NNN & ".xlsx")
   
            End If
       

       
    Workbooks(strNewWorkbookName).SaveAs Filename:=SAVRAN.Value & "\" & SheetNo
    strNewWorkbookName = ActiveWorkbook.Name
    Call �]�L(strNewWorkbookName, TABnumber)
    Workbooks(WWW).Save
    Application.CutCopyMode = False

    MsgBox "����ɏI�����܂���"
Exit Sub
ErrorHandler:
MsgBox ("�L�����Z������܂����B")
Workbooks(strNewWorkbookName).Close SaveCHANGES:=False
End Sub

Private Sub �폜(valListIndex As Variant)
    Dim lngLastRow As Long

    With Workbooks(lngOpenBookNumber).Worksheets(1)
        '�ŏI�s�擾
        lngLastRow = .Range("B65536").End(xlUp).Row
       
        .Activate
        
        If valListIndex = 0 Then
           
            
            '�Z���T�C�Y�̒���

            .Columns("B:L").EntireColumn.AutoFit
        
        ElseIf valListIndex = 2 Then
            'J�݌v�ύX
            .Range(Cells(1, 10), Cells(lngLastRow, 10)).Delete Shift:=xlShiftToLeft
            
             'I�s�����R�[�h
            .Range(Cells(1, 9), Cells(lngLastRow, 9)).Delete Shift:=xlShiftToLeft
            
            'G����
            .Range(Cells(1, 7), Cells(lngLastRow, 7)).Delete Shift:=xlShiftToLeft
            '�Z���T�C�Y�̒���
            .Columns("B:I").EntireColumn.AutoFit
       End If
    End With

End Sub

Private Sub OpenFile()

'�J��CSV�t�@�C���̎w����s��
On Error GoTo ErrorHandler
    strCsvbookName = Application.GetOpenFilename("CSV�t�@�C��(*.csv),*.csv")
    Workbooks.Open strCsvbookName
    strCsvbookName = Dir(strCsvbookName)
Exit Sub

'�L�����Z���ł��������̏���
ErrorHandler:
        'MsgBox "�I�����Ȃ������̂ŏI�����܂�"
        Exit Sub
End Sub

Private Sub �����ԍ�()

    Dim A As Long
    
    A = 1
    lngCellCounter = 1
    
    Do While Workbooks(strCsvbookName).Worksheets(1).Range("D" & A).Value <> ""
        Workbooks(strCsvbookName).Worksheets(1).Range("D" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
        A = A + 1
    Loop
    
    '�J��Ԃ��s���m��
    lngRowCounter = A
    
End Sub

Private Sub ���i��()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("H" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �����ҕi��()
    Dim A As Long
    Dim strTMP As String
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
'    For a = 1 To lngRowCounter
'        Workbooks(strCsvbookName).Worksheets(1).Range("J" & a).Copy Destination:= _
'        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(a, lngCellCounter)
'    Next

'20140414 ���l���Ƀ����N�������Ă�����A�u-���v�Ɣ����ҕi���̌��ɕt����

    For A = 1 To lngRowCounter
        strTMP = Workbooks(strCsvbookName).Worksheets(1).Range("J" & A).Value
        
        '���l�����󔒂ł͂Ȃ�������
        If Workbooks(strCsvbookName).Worksheets(1).Range("U" & A).Value <> "" Then
            strTMP = strTMP & "-" & Workbooks(strCsvbookName).Worksheets(1).Range("U" & A).Value
        End If
        
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter) = strTMP
        
    Next


End Sub

Private Sub �[���w����()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("L" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �[���w���()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("M" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �[������()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("P" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub ��n�ꏊ()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("V" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �s����()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AE" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �݌v�ύX�ԍ�()
    Dim A As Long
       
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AF" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �@��R�[�h()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AG" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub �������喼()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("AI" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next

End Sub

Private Sub KURODAUKEWATASHI()
    Dim A As Long
    
    '�ׂ̗�ֈړ�
    lngCellCounter = lngCellCounter + 1
    
    For A = 1 To lngRowCounter
        Workbooks(strCsvbookName).Worksheets(1).Range("FP" & A).Copy Destination:= _
        Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells(A, lngCellCounter)
    Next
 KA = Range(Cells(2, 3), Cells(Rows.Count, 3).End(xlUp)).Count
 Workbooks(strNewWorkbookName).Worksheets("Sheet1").Rows(1).AutoFilter FIELD:=7, Criteria1:="���޶����"
 Workbooks(strNewWorkbookName).Worksheets("Sheet1").Cells.Sort KEY1:=Worksheets("Sheet1").Range("C1"), _
 ORDER1:=xlAscending, Header:=xlYes
 KC = Range(Cells(2, 3), Cells(Rows.Count, 3).End(xlUp)).SpecialCells(xlCellTypeVisible).Count
End Sub
Private Sub ���ёւ����a()
    Dim lngLastRow As Long
    Dim i As Integer
    Dim i2 As Integer
    Dim strTmpText As String
    Dim lngTextNumber As Long
    Dim lngTmpRow  As Long
    Dim blnSpaceFlag As Boolean
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        
        '�󔒍s�t���O��������
        blnSpaceFlag = False
       
        '�ŏI�s�擾
        lngLastRow = .Range("A65536").End(xlUp).Row
       
        .Activate
       
        'C2�����ҕi�����ށ@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
              
        '�ŏI�s�܂ŉ�
        For i2 = 2 To lngLastRow
            'G2�ȉ��@��n�ꏊ���@���ï��@���@�ŏI�s+�󔒍s�̉���
            strTmpText = .Cells(i2, 7).Text
            '�܂܂�Ă�����A�ŏI�s+�󔒍s�̉���
            If InStr(strTmpText, "���ï�") Or InStr(strTmpText, "��� ���1 ���") > 0 Then
                '�󔒍s�������ĂȂ�����������
                If blnSpaceFlag = False Then
                   blnSpaceFlag = True
                   lngLastRow = lngLastRow + 1
                End If
                
                '�s���J�b�g���čŏI�s�̉��Ƀy�[�X�g
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '�󔒂܂ŗ����甲����
            If strTmpText = "" Then
                Exit For
            End If
        Next
    
        '�󔒍s�t���O��������
        blnSpaceFlag = False
       
        '�ŏI�s�܂ŉ�
        For i2 = 2 To lngLastRow
           
           '20160111 ��Ͳֳ���ޮ����ǉ�
            'G2�ȉ��@��n�ꏊ���@����ڼ�݁@���@�ŏI�s+�󔒍s�̉���
            strTmpText = .Cells(i2, 7).Text
            '�܂܂�Ă�����A�ŏI�s+�󔒍s�̉���
            If InStr(strTmpText, "����ڼ��") > 0 Or InStr(strTmpText, "��Ͳֳ����") > 0 Then
                '�󔒍s�������ĂȂ�����������
                If blnSpaceFlag = False Then
                    blnSpaceFlag = True
                    lngLastRow = lngLastRow + 1
                End If
               
                '�s���J�b�g���čŏI�s�̉��Ƀy�[�X�g
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '�󔒂܂ŗ����甲����
            If strTmpText = "" Then
                Exit For
            End If
        Next
       
        '�󔒍s�t���O��������
        blnSpaceFlag = False
        
        For i2 = 2 To lngLastRow
            '�����ҕi���R�[�h��
            strTmpText = Trim(.Cells(i2, 3).Text)
            lngTextNumber = Len(strTmpText)
               
            '5���� & "-" & 5�����̏ꍇ��
            If lngTextNumber = 11 Then
            
                '10�����ڂ�"-"�������牽�����Ȃ�
                If InStr(10, strTmpText, "-") <> 10 Then
                
                    strTmpText = Left(strTmpText, 1)
                       
                    '���̕�����"5"����Ȃ������ꍇ
                    If strTmpText <> "5" Then
                        '���̍s�폜
                        .Rows(i2).Delete
                        lngLastRow = lngLastRow - 1
                        i2 = i2 - 1
                    Else
                        '�󔒍s�������ĂȂ�����������
                        If blnSpaceFlag = False Then
                            blnSpaceFlag = True
                            lngLastRow = lngLastRow + 1
                        End If
                        
                        '�s���J�b�g���čŏI�s�̉��Ƀy�[�X�g
                        .Rows(i2).Cut
                        .Rows(lngLastRow + 1).Insert
                        i2 = i2 - 1
                       
                    End If
                End If
            End If
                    
            '�󔒂܂ŗ����甲����
            If strTmpText = "" Then
                Exit For
            End If
        Next
           
        '�󔒍s�t���O��������
        blnSpaceFlag = False
        
        '�ŏI�s�܂ŉ�
        For i2 = 2 To lngLastRow
           
            'G2�ȉ��@��n�ꏊ���@CKD�@���@�ŏI�s+�󔒍s�̉���
            strTmpText = .Cells(i2, 7).Text
            '�܂܂�Ă�����A�ŏI�s+�󔒍s�̉���
            If InStr(strTmpText, "CKD") > 0 Then
                '�󔒍s�������ĂȂ�����������
                If blnSpaceFlag = False Then
                    blnSpaceFlag = True
                    lngLastRow = lngLastRow + 1
                End If
               
                '�s���J�b�g���čŏI�s�̉��Ƀy�[�X�g
                .Rows(i2).Cut
                .Rows(lngLastRow + 1).Insert
                i2 = i2 - 1
            End If
            
            '�󔒂܂ŗ����甲����
            If strTmpText = "" Then
                Exit For
            End If
            
        Next
        
        
        '�ォ��󔒍s���o�Ă���܂ł̊Ԃ��AF2�ȉ��@�[������1�@�����ŕ��ёւ�
        lngTmpRow = .Range("A1").End(xlDown).Row
        
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 6), ORDER1:=xlAscending
        
        'A2�����ԍ��@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 1), ORDER1:=xlAscending
        
        'C2�����ҕi�����ށ@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngTmpRow, 11)).Sort _
            KEY1:=.Cells(2, 3), ORDER1:=xlAscending
        
        'A���1��}��
        .Columns(1).Insert
            
        '�Z���T�C�Y�̒���
        .Columns("B:L").EntireColumn.AutoFit
        
    End With
End Sub

Private Sub ���ёւ�SKK()
    Dim lngLastRow As Long
    Dim i As Integer
    Dim strTmpText As String
    Dim blnSpaceFlag As Boolean
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        
        '�󔒍s�t���O��������
        blnSpaceFlag = False
       
        '�ŏI�s�擾
        lngLastRow = .Range("A65536").End(xlUp).Row
       
        .Activate
       
        'C2�����ҕi�����ށ@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
                
        'J2�@���ށ@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngLastRow, 11)) _
                .Sort KEY1:=.Cells(2, 10), ORDER1:=xlAscending
                
        '�ŏI�s�܂ŉ�
        For i = 2 To lngLastRow
            '�������喼���󂾂�����폜
            strTmpText = .Cells(i, 11).Text
            If strTmpText <> "" Then
                blnSpaceFlag = True
            End If
        Next
        
        If blnSpaceFlag = False Then
            .Columns(11).Delete
        End If
        
        '�Z���T�C�Y�̒���
        .Columns("A:K").EntireColumn.AutoFit
                
    End With
    
End Sub

Private Sub ���ёւ����̑�()
    Dim lngLastRow As Long
    Dim i2 As Integer
    
    With Workbooks(lngOpenBookNumber).Worksheets(1)
        '�ŏI�s�擾
        lngLastRow = .Range("A65536").End(xlUp).Row
        
        'C2�����ҕi�����ށ@�����ŕ��ёւ�
        .Range(Cells(2, 1), Cells(lngLastRow, 8)) _
                .Sort KEY1:=.Cells(2, 3), ORDER1:=xlAscending
             
                
        '�Z���T�C�Y�̒���
        .Columns("A:H").EntireColumn.AutoFit
    End With
   
End Sub


Private Sub KEISEN(ByVal SPNAME As String) '���C�����Ɍr��������
Dim AHA As Range
Set AHA = Cells(Rows.Count, 4).End(xlUp)

Dim AST, STR, LineSt As Integer
LineSt = xlContinuous
AST = ActiveSheet.Cells.Find("�@����").Offset(1, 0).Column
    If UserForm1.ComboBox1.Text = "SKK" Then
    STR = 0
    ElseIf UserForm1.ComboBox1.Text = "���a�V�[��" Then
    STR = 1
    Else
    STR = -3
    End If
Dim PPPC As Range
    For Each PPPC In Range(Cells(2, AST), Cells(AHA.Row, AST))
        If PPPC.Row = AHA.Row Then Exit Sub
        If PPPC.Offset(1, -6).Value = "" Then
        GoTo NNN
        End If
        If PPPC.Value = PPPC.Offset(1, 0).Value Then 'ײ݂��ς��Ȃ���Δ����r��������
           
        Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).ColorIndex = 15
       
        GoTo NNN
        Else
            If PPPC.Offset(1, 0).Value <> PPPC.Value Then 'ײ݂��ς������r��������
            
            Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range(Cells(PPPC.Row, 1), Cells(PPPC.Row, 18 + STR)).Borders(xlEdgeBottom).ColorIndex = 56
            
            End If
        End If
    
NNN:
    
    Next PPPC
End Sub

Sub KDODNA() '���y�[�W

Dim KRANGE As Range
Set KRANGE = Range("D1").End(xlDown).Offset(2, 0)
ActiveSheet.HPageBreaks.Add BEFORE:=KRANGE
End Sub



Sub ���c����() '!CSV�̐������������s��!
Dim A, D, E As Variant
Dim WC, R, C, P As Integer
Dim KURODACSV As New Collection
WC = 0
E = 1

A = Application.GetOpenFilename("CSV�t�@�C��(*.csv),*.csv", Title:="Ctrl�������Ȃ���csv�ް���2�ȏ�I�����Ă��������B", MultiSelect:=True)
    
    For Each D In A
    Workbooks.Open D
    KURODACSV.Add Item:=ActiveWorkbook
         WC = WC + 1
    Next
    
    Do While E <= UBound(A) '���[�N�u�b�N�̐������J��Ԃ��B
        If E <> 1 Then
            KURODACSV(E).Activate
            If KURODACSV(E).Worksheets(1).Range("B2").Offset(1, 0).Value = "" Then
            KURODACSV(E).Worksheets(1).Range(Cells(2, 1), Cells(2, Columns.Count).End(xlToLeft)).Copy
            Else
            R = KURODACSV(E).Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row '�ŏI�s
            C = KURODACSV(E).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '�ŏI��
            KURODACSV(E).Worksheets(1).Range(Cells(2, 1), Cells(R, C)).Copy
            End If
        Else
        GoTo SSS
        End If
        
        KURODACSV(1).Worksheets(1).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial
        Application.CutCopyMode = False
        KURODACSV(E).Close SaveCHANGES:=False

SSS:
        E = E + 1
    Loop
    strCsvbookName = KURODACSV(1).Name
    
    
End Sub

Sub KURODAKEISEN()
Dim KURO As Range
Dim KKR, KKC, LASTKR As Integer
LASTKR = Cells(Rows.Count, 1).End(xlUp).Row
    For Each KURO In Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp))
        If KURO.Row = LASTKR Then
        Exit Sub
        End If
    KKR = KURO.EntireRow.Row
    KKC = KURO.EntireColumn.Column
    Range(Cells(KKR, KKC), Cells(KKR, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(KKR, KKC), Cells(KKR, 15)).Borders(xlEdgeBottom).ColorIndex = 15
    Next
End Sub


