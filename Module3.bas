Attribute VB_Name = "Module3"

Sub messagecompleted()
 Dim Msg, Buttons, Title
Msg = "Task completed"
Buttons = vbOKOnly
Title = "JSON converter"
MsgBoxCompleted = MsgBox(Msg, Buttons, Title)

End Sub

