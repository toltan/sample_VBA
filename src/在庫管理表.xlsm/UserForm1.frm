VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "セル範囲選択"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If RefEdit1.Value = "" Then
MsgBox "セル範囲が選択されていません。"
Exit Sub
Else
Dim SYUKKO As Range
Set SYUKKO = Range(Me.RefEdit1.Value)
End If

Unload UserForm1
Workbooks.Open ("C:\Users\owner\Desktop\重慶\重慶出庫明細")
Workbooks("重慶出庫明細.xlsx").Worksheets(1).Range(Cells(7, 1), Cells(Rows.Count, Columns.Count)).Delete
SYUKKO.Copy
Workbooks("重慶出庫明細.xlsx").Worksheets(1).Range("B7").PasteSpecial Paste:=xlAll
If Range("I7").Offset(1, 0).Value = ("") Then
With Workbooks("重慶出庫明細.xlsx").Worksheets(1)
.Range("D7:F100").Delete
.Range("F7:I100").Delete
.Range("H7:EE100").Delete
End With
Else
With Workbooks("重慶出庫明細.xlsx").Worksheets(1)
.Range("H7").End(xlDown).Offset(1, 0).Select
.Rows(ActiveCell.Row & ":" & 100).Delete
.Range("D7:G100").Delete
.Range("G7:I100").Delete
.Range("H7:EE100").Delete
End With
End If
Workbooks("重慶出庫明細.xlsx").Worksheets(1).Activate
End Sub


Private Sub UserForm_Click()

End Sub
