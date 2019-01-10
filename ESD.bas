Attribute VB_Name = "ESD"
Sub ESDsheet()

Rows("1:1").Select
range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets.Add After:=activesheet
activesheet.Paste
range("L1").Select
ActiveCell.FormulaR1C1 = "Account Owner's email"
Selection.Font.Bold = True
range("M1").Select
ActiveCell.FormulaR1C1 = "Account Owner's Name"
Selection.Font.Bold = True
range("N1").Select
ActiveCell.FormulaR1C1 = "Contact Person's Name"
Selection.Font.Bold = True
range("B1").Select
ActiveCell.FormulaR1C1 = "Amount"
Selection.Font.Bold = True
range("O1").Select
ActiveCell.FormulaR1C1 = "Contact's Account"
Selection.Font.Bold = True

Columns("A:O").EntireColumn.AutoFit
ActiveWindow.Zoom = 85
Columns("B:B").Select
Selection.EntireColumn.Hidden = True
Columns("C:C").Select
Selection.EntireColumn.Hidden = True
Columns("D:D").Select
Selection.EntireColumn.Hidden = True
Columns("I:I").Select
Selection.EntireColumn.Hidden = True
Columns("J:J").Select
Selection.EntireColumn.Hidden = True
Columns("E:E").ColumnWidth = 10
Columns("K:K").ColumnWidth = 5.86


End Sub
