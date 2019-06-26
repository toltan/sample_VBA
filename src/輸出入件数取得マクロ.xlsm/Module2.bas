Attribute VB_Name = "Module2"
Option Explicit

Sub hjk()

Dim underTableCol As Integer
Dim totalRow As Integer

underTableCol = ActiveSheet.Cells.Find(what:="—A“üŽÒ", lookat:=xlWhole, after:=ActiveSheet.Cells.Find(what:="—A“üŽÒ", lookat:=xlWhole)).Column
totalRow = ActiveSheet.Cells.Find(what:="TOTAL", lookat:=xlWhole)
Cells(51, underTableCol).Select

End Sub
