Attribute VB_Name = "Module1"
Public Sub jsonConverter(engine As String, engineRow As Integer)

'Optimization START
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Debug.Print ("enginerow is " & engineRow)
'Code for Excel to Json START
Dim rng2 As Range, items As New Collection, myitem As New Dictionary, i As Integer, cell As Variant
Dim rng1 As Range
Debug.Print ("A" & engineRow & ":XFD" & engineRow)
Set rng1 = Range("A" & engineRow & ":XFD" & engineRow)
Dim os As Integer

os = 0
If (engine = "" Or engine = "Engine" Or engine = "Select engine") Then
Debug.Print ("Didn't select anything")
MsgBox4 = MsgBox("Please select a Engine", vbOKOnly, "Warning")
Else
For Each cell In rng1
    If (cell.Value = engine) Then
        Debug.Print ("SUCCES")
        Exit For
    End If
os = os + 1
Next
Debug.Print ("os = " & os)

Debug.Print ("A" & engineRow & ":A30000")
Set rng2 = Range("A" & engineRow & ":A30000")
'Set rng = Range(Sheets(2).Range("A2"), Sheets(2).Range("A2").End(xlDown)) use this for dynamic range
i = 0
    For Each cell In rng2
        If Not (IsEmpty(cell.Offset(0, os).Value) Or Application.IsNA(cell.Offset(0, 14).Value) Or CStr(cell.Offset(0, 14).Value) = "#N/A") Then
         
myitem(cell.Value) = cell.Offset(0, os).Value
items.Add myitem
Set myitem = Nothing
        End If
i = i + 1
Next
Sheets(1).Range("A30000").Value = ConvertToJson(items, Whitespace:=2)
'Code for Excel to Json END

'Optimization END
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Call printToFile(engine)
Call removeText
Call messagecompleted
End If
End Sub
