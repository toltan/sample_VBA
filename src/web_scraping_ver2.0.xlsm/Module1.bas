Attribute VB_Name = "Module1"
Sub SortFill1() '差枚数、最終ゲーム数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("差枚数", "最終ゲーム数"), Operator:=xlFilterValues
End Sub

Sub SortFill2() '総回転数、BIG回数、REG回数、ART初当たりでソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("総回転数", "BIG回数", "REG回数", "ART初当たり回数"), Operator:=xlFilterValues
End Sub

Public Sub SortFill3() '差枚数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:="差枚数"
End Sub

Public Sub SortFill4() '総回転数、BIG回数、REG回数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("総回転数", "BIG回数", "REG回数"), Operator:=xlFilterValues
End Sub

Public Sub SortFill5() 'ART初当たり回数、差枚数、最終ゲーム数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("ART初当たり回数", "差枚数", "最終ゲーム数"), Operator:=xlFilterValues
End Sub

Public Sub SortFill6() 'ART初当たり回数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("ART初当たり回数"), Operator:=xlFilterValues
End Sub

Public Sub SortFill7() '最終ゲーム数でソート
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:=Array("最終ゲーム数"), Operator:=xlFilterValues
End Sub

Public Sub SortFillClear() 'フィルタークリア
ActiveSheet.Range("A6").CurrentRegion.AutoFilter _
FIELD:=3, Criteria1:="<>("")", Operator:=xlFilterValues
End Sub
