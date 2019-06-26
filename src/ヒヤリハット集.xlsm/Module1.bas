Attribute VB_Name = "Module1"
Option Explicit
Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As LongPtr


Public Sub inputData()
Dim myname As String
Dim num As Long
Dim rtn As LongPtr

myname = String(250, Chr(0))
num = Len(myname)
rtn = GetUserName(myname, num)
MsgBox myname
End Sub

Function LoginName() As String
    LoginName = CreateObject("WScript.Network").UserName
End Function

Function ExcelUserName() As String
    ExcelUserName = Application.UserName
End Function
