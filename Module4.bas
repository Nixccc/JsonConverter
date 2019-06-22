Attribute VB_Name = "removetext"
Public Sub removeText()
Range("A30000").Select
ActiveCell.FormulaR1C1 = ""
Range("A1").Select
End Sub

