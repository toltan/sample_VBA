Attribute VB_Name = "Module2"
Sub PRIN()
Dim PPP As String
PPP = Application.ActivePrinter
MsgBox (PPP)
Application.ActivePrinter = "EPSONAB0011 (PX-204) on Ne04:"
MsgBox (Application.ActivePrinter)
Application.ActivePrinter = PPP
MsgBox (Application.ActivePrinter)

End Sub

Sub PFP()
Worksheets(1).Range("S2").Value = ActivePrinter
End Sub
