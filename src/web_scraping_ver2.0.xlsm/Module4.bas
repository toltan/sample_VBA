Attribute VB_Name = "Module4"
Public Sub fn() 'find���g���ꍇ�́A�����G���A���ڍׂɎw�肵�Ȃ��Ƃ��������Ȃ邱�Ƃ����肤��icells�j�̑��l�͔�����
Dim a, b, c, d, e, datavalue, ro, co, rrr As Long
Dim datano, dataname, dayval As String
Dim fc, wsh As Object
Dim ccc As Range

dayval = UserForm1.ComboBox1.Text '���t�I��

Application.ScreenUpdating = False '�����\�����Ȃ�
Worksheets(cb1t).Activate '�]�L���錎�̃V�[�g���A�N�e�B�u��

If Worksheets(cb1t).AutoFilterMode = False Then Worksheets(cb1t).Rows(6).AutoFilter

Set ccc = Worksheets(cb1t).Range(Cells(7, 1), Cells(Rows.Count, 1).End(xlUp))
Worksheets("scraiping").Select
Worksheets("scraiping").Range("B1").Select
dataname = Selection.Value
datavalue = 1
e = 0

    Do While Selection.Value <> ""
        a = Worksheets("scraiping").Cells(datavalue + 1, 2).Row
        b = Worksheets("scraiping").Cells(datavalue + 1, 2).Offset(5, 2).Row
        c = 2 '��̗�ԍ�
    
        Do While Selection.Value <> ""
            Cells(a, c).Select
            
            If Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "��ԍ�" Then '
                datano = Cells(a, c).Value
                Set fc = Worksheets(cb1t).Columns("A").Find(what:=datano, LookIn:=xlValues, LookAt:=xlWhole)
                    
                    If fc Is Nothing Then
                        Worksheets(cb1t).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Value = datano
                        Worksheets(cb1t).Cells(Rows.Count, 1).End(xlUp).Offset(0, 1).Value = Trim(dataname) '
                        Set fc = Worksheets(cb1t).Columns("A").Find(what:=datano, LookIn:=xlValues, LookAt:=xlWhole)
                        Worksheets(cb1t).Activate
                        Set ccc = Worksheets(cb1t).Range(Cells(7, 1), Cells(Rows.Count, 1).End(xlUp))
                    ElseIf Trim(fc.Offset(0, 1).Value) <> Trim(dataname) Then
                        
                        For Each ron In ccc
                            If ron.Value = datano Then
                                If Trim(ron.Offset(0, 1).Value) = Trim(dataname) Then
                                GoTo cont
                                End If '
                            End If
                        Next
                        '�Y�����Ȃ�������V���ɗ������
                        Worksheets(cb1t).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Value = datano
                        Worksheets(cb1t).Cells(Rows.Count, 1).End(xlUp).Offset(0, 1).Value = Trim(dataname)
                        Set fc = Worksheets(cb1t).Columns("A").Find(what:=datano, LookIn:=xlValues, LookAt:=xlWhole)
                        Worksheets(cb1t).Activate
                        Set ccc = Worksheets(cb1t).Range(Cells(7, 1), Cells(Rows.Count, 1).End(xlUp))
                    End If
cont:
                Worksheets(cb1t).Activate
                
                For Each ran In ccc
                
                    If ran.Value = datano Then
                    
                        If Trim(ran.Offset(0, 1).Value) = Trim(dataname) Then
                            ro = ran.Row '�Y������s���擾�A������玟�̍s��
                            Exit For
                        Else
                            GoTo continue
                        End If
                        
                    Else
                        GoTo continue
                    End If
continue:
                Next
                    
                Worksheets("scraiping").Activate
                
            ElseIf Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "����]��" Then
                co = Worksheets(cb1t).Rows(6).Find(what:=dayval, LookIn:=xlValues, LookAt:=xlWhole).Column '�]�L����̗���擾
            
                If e = 0 Then
                    e = e + 1
                    
                    If Worksheets(cb1t).Cells(ro, co).Value <> "" Then '�]�L��Ɋ��Ƀf�[�^���������ꍇ�̏���
                        d = MsgBox("���ɓ��͂���Ă���f�[�^������܂��B" & vbCrLf & "�㏑�����܂����H", vbYesNo)
                        
                        If d = vbNo Then
                            MsgBox ("���f���܂����B")
                            Worksheets(cb1t).Rows(6).AutoFilter
                            Unload UserForm1
                            Exit Sub
                        End If
                        
                    End If
                    
                End If
                
            Worksheets(cb1t).Cells(ro, co).Value = Worksheets("scraiping").Cells(a, c).Value
            
            ElseIf Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "BIG��" Then
                Worksheets(cb1t).Cells(ro + 1, co).Value = Worksheets("scraiping").Cells(a, c).Value
            ElseIf Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "REG��" Then
                Worksheets(cb1t).Cells(ro + 2, co).Value = Worksheets("scraiping").Cells(a, c).Value '
            ElseIf Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "�ŏI�Q�[����" Then
                Worksheets(cb1t).Cells(ro + 6, co).Value = Worksheets("scraiping").Cells(a, c).Value
            ElseIf Worksheets("scraiping").Cells(a, c).End(xlToLeft).Value = "������" Then
                Worksheets(cb1t).Cells(ro + 5, co).Value = Worksheets("scraiping").Cells(a, c).Value
            End If
            
            If a = b Then
                c = c + 1
                a = Worksheets("scraiping").Cells(datavalue + 1, 2).Row
            Else '�ŏI�s�łȂ���Ύ��̃f�[�^��
                a = a + 1
            End If
            
        Loop
        
        datavalue = datavalue + 7
        Worksheets("scraiping").Select
        Worksheets("scraiping").Cells(datavalue, 2).Select
        dataname = Selection.Value

    Loop
    
Worksheets(cb1t).Rows(6).AutoFilter
Unload UserForm1
Application.ScreenUpdating = True
    
'�|�b�v�A�b�v
Set wsh = CreateObject("WScript.Shell")
wsh.popup Text:="complete!!", secondstowait:=1, Type:=vbInformation
Set wsh = Nothing

End Sub
