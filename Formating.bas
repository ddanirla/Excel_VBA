Attribute VB_Name = "Module1"
Sub vidas_formatting()
Attribute vidas_formatting.VB_Description = "formate vidas data to validate data formating"
Attribute vidas_formatting.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' vidas_formatting Macro
' formate vidas data to validate data formating
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.NumberFormat = "dd/mm/yyyy;@"
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.NumberFormat = "0"
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.NumberFormat = "0"
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.NumberFormat = "@"
    ActiveCell.Offset(1, 0).Range("Table2[[#Headers],[s/n ]]").Select
End Sub
