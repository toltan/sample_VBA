Attribute VB_Name = "Module1"
Sub SortFill1() '�������A�ŏI�Q�[�����Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("������", "�ŏI�Q�[����"), Operator:=xlFilterValues
End Sub

Sub SortFill2() '����]���ABIG�񐔁AREG�񐔁AART��������Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("����]��", "BIG��", "REG��", "ART���������"), Operator:=xlFilterValues
End Sub

Public Sub SortFill3() '�������Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:="������"
End Sub

Public Sub SortFill4() '����]���ABIG�񐔁AREG�񐔂Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("����]��", "BIG��", "REG��"), Operator:=xlFilterValues
End Sub

Public Sub SortFill5() 'ART��������񐔁A�������A�ŏI�Q�[�����Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("ART���������", "������", "�ŏI�Q�[����"), Operator:=xlFilterValues
End Sub

Public Sub SortFill6() 'ART��������񐔂Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("ART���������"), Operator:=xlFilterValues
End Sub

Public Sub SortFill7() '�ŏI�Q�[�����Ń\�[�g
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("�ŏI�Q�[����"), Operator:=xlFilterValues
End Sub

Public Sub SortFillClear() '�t�B���^�[�N���A
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:="<>("")", Operator:=xlFilterValues
End Sub
