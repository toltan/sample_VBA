Attribute VB_Name = "Module3"
Sub MAEBASI()
Dim MAEOG As Range
Dim A As Integer

For Each MAEOG In Range(Cells(Cells.Find(WHAT:="�����ҕi������-���l").Offset(1, 0).Row, 4), Cells(Rows.Count, 4).End(xlUp))
MAEOG.Select
If MAEOG.Value Like "*-0290" Or MAEOG.Value Like "*-0291" Or MAEOG.Offset(0, 4).Value = "KF" Then '�������-0290,����-0291���܂܂�Ă����ꍇ�̔���
MAEOG.EntireRow.Font.Strikethrough = True
End If
Next
End Sub

Sub KIRYU()
Dim KIR As Range
Dim A As Variant
Dim R As Integer
Dim C As Integer
Dim TARGET As String
   
    TARGET = "�����ҕi������-���l"
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    A = Array("1013-1400", "1410-1400", "7466-1400")
    
    '�i�Ԍ���
For Each KIR In Range(Cells(2, C), Cells(Cells(Rows.Count, 2).End(xlUp).Row, C))
    
    If KIR.Value = A(0) Or KIR.Value = A(1) Or KIR.Value = A(2) Then '�z��̂����ꂩ�ƈ�v�����ꍇ�̔���
        
        If Workbooks(WWW).Worksheets("DATABASE") _
        .Cells.Find(WHAT:=KIR).Offset(0, 2) = "�Z" Then 'DATABASE�̒����|�������Z��������i�Ԃ�-����t����
        KIR.Value = KIR & "��"
        End If
    ElseIf KIR.Value Like "*-*-*" Or KIR.Offset(0, 3).Value = "���� CKD" Or KIR.Offset(0, 4).Value = "KF" Then
    Rows(KIR.Row).Select
    Selection.Font.Strikethrough = True
    
    End If
    
Next
End Sub


Sub SKKSort()
Dim R As Integer
Dim C As Integer
Dim TARGET As String
    
    If UserForm1.ComboBox1.Text = "SKK" Then Columns("A").Insert Shift:=xlToRight
       
    TARGET = "�����ҕi������-���l"
    
    For AAA = 1 To 2
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    
    ActiveSheet.Range("A1").Sort KEY1:=ActiveSheet.Cells(R, C), ORDER1:=xlAscending, Header:=xlGuess
    
    TARGET = "�@����"

    Next


End Sub

Sub SKKEREMA() '�G���}��ʃy�[�W�Ɉړ�


Dim R As Integer
Dim C As Integer
Dim TARGET As String

    TARGET = "��n�ꏊ��"
    R = ActiveSheet.Cells.Find(WHAT:=TARGET).Row
    C = ActiveSheet.Cells.Find(WHAT:=TARGET).Column
    
For Each EREMA In Range(Cells(2, C), Cells(Cells(Rows.Count, C).End(xlUp).Row, C))
EREMA.Select '��ŏ���
    If EREMA.Value = "���ï� ���" Then
    Rows(EREMA.Row).Select
        
        If EREMA.Offset(1, 0).Value <> "" And EREMA.Offset(1, 0).Value <> "���ï� ���" _
        And EREMA.End(xlDown).Value <> "���ï� ���" Then
        '�ŏI�s�łȂ��A1�i�����G���}����Ȃ��A�ŏI�s���G���}����Ȃ�������ŏI�s��
        Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "���ï� ���" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA '���y�[�W
            End If
            ElseIf EREMA.Offset(1, 0).Value <> "" And EREMA.Offset(1, 0).Value <> "���ï� ���" _
         Then '�ŏI�s�łȂ��A1�i�����G���}����Ȃ�������ŏI�s��
         Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "���ï� ���" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
            End If
        ElseIf EREMA.Offset(1, 0).Value <> "" And EREMA.End(xlDown).Value <> "���ï� ���" _
         Then '�ŏI�s�łȂ��A�ŏI�s���G���}����Ȃ�������ŏI�s��
         Selection.Cut Destination:=Rows(Cells(Rows.Count, C).End(xlUp).Offset(1, 0).Row)
            If Cells(Rows.Count, C).End(xlUp).Offset(-1, 0).Value <> "���ï� ���" Then
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
            End If
        ElseIf EREMA.Offset(1, 0).Value = "" And EREMA.Offset(-1, 0).Value <> "���ï� ���" _
         Then '�ŏI�s�ŁA1�i�オ�G���}����Ȃ���������y�[�W�B1���������Ȃ������牽�����Ȃ�
            If EREMA.Offset(-1, 0).Value <> "��n�ꏊ��" Then ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
        ElseIf EREMA.Offset(1, 0).Value = "���ï� ���" And EREMA.End(xlDown).Value = "���ï� ���" And EREMA.Offset(-1, 0).Value <> "���ï� ���" Then
        '1�i�����G���}�ŁA1�i�オ�G���}�łȂ��A�ŏI�s���G���}�̏ꍇ���y�[�W
            ActiveSheet.HPageBreaks.Add BEFORE:=EREMA
        End If
           
    ElseIf EREMA.Value = "��ܺ�ï��" Or EREMA.Value = "��İ �����" Or EREMA.Value = "̺� �����" Or EREMA.Value = "���ï� ���" Then
        Rows(EREMA.Row).Select
        Selection.Font.Strikethrough = True

    
        
            
    End If
Next EREMA


'�󔒍폜
For Each EREMA In Range(Cells(2, C), Cells(Cells(Rows.Count, C).End(xlUp).Row, C))
    EREMA.Select
    If EREMA.Value = "" Then
    Rows(EREMA.Row).Delete '2�i�󔒂������Ă���ꍇ1�i�����̂ŁA�΍�
    
    End If
Next EREMA

End Sub

