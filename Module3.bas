Attribute VB_Name = "complete"

Sub messagecompleted()
 Dim Msg, Buttons, Title
Msg = "Task completed"
Buttons = vbOKOnly
Title = "JSON converter"
MsgBoxCompleted = MsgBox(Msg, Buttons, Title)

End Sub

