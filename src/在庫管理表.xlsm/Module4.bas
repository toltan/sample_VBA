Attribute VB_Name = "Module4"
Option Explicit

Sub ピボット更新()
Attribute ピボット更新.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' ピポット更新 Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Sheets("ピボット").Select
    ActiveSheet.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotSelect "", xlDataAndLabel, True
    ActiveSheet.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotCache.Refresh
    Worksheets("在庫一覧").Select
    
End Sub
