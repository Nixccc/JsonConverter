Attribute VB_Name = "Module4"
Public Sub removeText()
Range("A30000").Select
ActiveCell.FormulaR1C1 = ""
Range("A1").Select
End Sub

