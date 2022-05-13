Attribute VB_Name = "Module1"
Sub Get_test_Results()
Attribute Get_test_Results.VB_Description = "In this instance i want to get the text ""POS"" and ""NEG"" \nfrom a test results data  so much so that i wont have to to this boring task over and over again ....... i would just need to call its shortcut"
Attribute Get_test_Results.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' Get_test_Results Macro
' In this instance i want to get the text "POS" and "NEG"  from a test results data  so much so that i wont have to to this boring dask over and over again ....... i would just need to call its shortcut
'
' Keyboard Shortcut: Ctrl+Shift+G
'
    ActiveCell.Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=LEFT([@TDR],3)"
End Sub
