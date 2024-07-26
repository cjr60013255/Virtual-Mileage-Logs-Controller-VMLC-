Attribute VB_Name = "Area_Email_From_Roster"
Sub Area_Email_From_Roster()
Attribute Area_Email_From_Roster.VB_ProcData.VB_Invoke_Func = " Area_Email_From_Roster"
'
' Area_Email_From_Roster Macro
'

'
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1:AA2").Select
    Selection.EntireRow.Delete
    ActiveWindow.SmallScroll Down:=-300
    Range("A1").Select
    
    Columns("A:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:B").Select
    Selection.ColumnWidth = 26.14
    
    Columns("A:B").Select
    ActiveSheet.Range("$A$1:$B$226").removeduplicates Columns:=Array(1, 2), _
        Header:=xlYes
    Columns("A:A").Select
    Selection.Cut
    Columns("C:C").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "office"
    Columns("C:IV").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Sub
