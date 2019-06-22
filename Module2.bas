Attribute VB_Name = "Module2"
Public Sub printToFile(enginename As String)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim wk_master As Workbook
Dim ws_master As Worksheet
Dim cell_value As String
Dim name_to_save As String
Dim cell As Variant
Set wk_master = ActiveWorkbook
Set ws_master = wk_master.Worksheets(1) '

name_to_save = enginename

cell_value = ws_master.Cells(30000, "A").Value

Set fso = CreateObject("Scripting.FileSystemObject")
Dim Fileout As Object
Set Fileout = fso.CreateTextFile(ThisWorkbook.Path & "\Json\" & name_to_save & ".json", True, True)
Fileout.Write cell_value
Fileout.Close
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
