Attribute VB_Name = "Module3"
Public cb1t As String
Public Sub osjs()

cb1t = Worksheets("scraiping").ComboBox1.Text
Worksheets(cb1t).Activate
UserForm1.Show (vbModeless)

End Sub
