Attribute VB_Name = "Module2"
Sub PRINTSET()
Attribute PRINTSET.VB_ProcData.VB_Invoke_Func = "P\n14"
Dim WS As Worksheet
For Each WS In Worksheets
WS.PageSetup.PaperSize = 156
WS.PageSetup.Orientation = xlPortrait
Next
ActiveWorkbook.Save
End Sub

Sub JJJ()
Dim BUF, MSG As String
Dim WS As Worksheet
With CreateObject("WScript.Shell")
        .CurrentDirectory = "\\192.168.1.230\�ɐ��肌�Z���^�[\�k�Z���^�[�p�`�["
    End With

MSG = Dir("*", vbDirectory)
    Do While MSG <> ""
    If MSG <> "." And MSG <> ".." Then
     With CreateObject("WScript.Shell")
        .CurrentDirectory = "\\192.168.1.230\�ɐ��肌�Z���^�[\�k�Z���^�[�p�`�[\" & "�n�s" '�����ŕҏW
        
    End With
    BUF = Dir("*.x*")
   
            Do While BUF <> ""
            Workbooks.Open (BUF)
            
            For Each WS In Worksheets
            WS.PageSetup.PaperSize = 129
            WS.PageSetup.Orientation = xlPortrait
            Next
            ActiveWorkbook.Save
            
            
            
            
            
            
            
            
            
            Workbooks(BUF).Close
            BUF = Dir()
            Loop
    
    End If
   
   
    Loop
MsgBox ("COMPLETE!"), vbInformation
End Sub

