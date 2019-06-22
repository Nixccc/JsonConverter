VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mouse_down As Boolean
Dim mouse_starting_X As Double
Dim mouse_starting_Y As Double
Private Sub CommandButton1_Click()
Call filereader.jsonConverter(ComboBox1.Value)
End Sub


Private Sub CommandButton2_Click()
UserForm1.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
mouse_down = False
HideTitleBar Me
UserForm1.Font.Size = 10
UserForm1.Font.Charset = Calibri
UserForm1.BackColor = &HFFFFFF
UserForm1.BorderColor = &HA9A9A9
ComboBox1.AddItem ("Select engine")
ComboBox1.ListIndex = 0
ComboBox1.BorderStyle = fmBorderStyleSingle
ComboBox1.BorderColor = &HA9A9A9
CommandButton1.BackColor = RGB(255, 255, 255)
CommandButton2.BackColor = RGB(255, 255, 255)


Dim rng As Range, cell As Variant
Set rng = Range("C4:XFD4")
For Each cell In rng
    If Not (IsEmpty(cell.Value) Or Application.IsNA(cell.Value) Or CStr(cell.Value) = "#N/A") Then
        ComboBox1.AddItem cell.Value
        Debug.Print (cell.Value)
        End If
Next
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
mouse_down = True
mouse_starting_X = X
mouse_starting_Y = Y
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If mouse_down = True Then
    UserForm1.Left = UserForm1.Left + (X - mouse_starting_X)
    UserForm1.Top = UserForm1.Top + (Y - mouse_starting_Y)
End If

End Sub
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
mouse_down = False
End Sub

Private Sub ComboBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetComboBoxHook(ComboBox1)
End Sub

Private Sub ComboBox1_LostFocus()
    Call RemoveComboBoxHook
End Sub
