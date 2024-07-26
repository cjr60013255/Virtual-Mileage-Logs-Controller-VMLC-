Attribute VB_Name = "ConvertVehiclesAssignmentReport"
Sub Convert_Vehicles_Assignment_Report_For_VMLC()
Attribute ConvertVehiclesAssignmentReportForVMLC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ConvertVehiclesAssignemntReportForVMLC Macro
'

'
    Range("A1").Select
    ActiveCell.Rows("1:2").EntireRow.Select
    Selection.Delete Shift:=xlUp
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1:L4").Select
    Selection.EntireRow.Delete
    
    Columns("D:L").Select
    Selection.Copy
    Columns("E:E").Select
    ActiveSheet.Paste
    Columns("D:D").Select
    Selection.ClearContents
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Area 2"
    Range("D2").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Designated Driver"
    Range("M1").Select
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    ActiveCell.FormulaR1C1 = "Designated Driver 2"
    Cells.Replace what:="" & Chr(10) & "", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, _
        FormulaVersion:=xlReplaceFormula2
    Cells.Replace what:=" - Can Drive - Designated Driver", Replacement:="_", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace what:="_*", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
    Cells.Replace what:="Cannot", Replacement:="Can", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace what:="*Can Drive", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Columns("M:M").ColumnWidth = 24.86
    Columns("M:N").Select
    Selection.ColumnWidth = 25.14
    Columns("K:K").Select
    Selection.ColumnWidth = 15.29
    Columns("I:I").Select
    Selection.ColumnWidth = 20.29
    Columns("H:H").ColumnWidth = 13.86
    Columns("G:G").Select
    Selection.ColumnWidth = 12.57
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("C:D").Select
    Range("D1").Activate
    Selection.ColumnWidth = 10
    Columns("A:B").Select
    Range("B1").Activate
    Selection.ColumnWidth = 12.57
    Range("A1").Select
    
    Columns("F:F").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("F1:F83") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:N83")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F1").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add(Range("F2:F79"), _
        xlSortOnFontColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(156, 0 _
        , 6)
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:N79")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    fTwo = Range("F2").value
    fThree = Range("F3").value
    Do While fTwo = fThree
        
        
        Range("C3").Select
        Selection.Cut
        Range("D2").Select
        ActiveSheet.Paste
        Range("M3").Select
        Selection.Cut
        Range("N2").Select
        ActiveSheet.Paste
        Range("K2").Select
        Selection.Copy
        Range("J3").Select
        ActiveSheet.Paste
        Range("I3").Select
        Application.CutCopyMode = False
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]+RC[2]"
        Range("I3").Select
        Selection.Copy
        Range("K2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Rows("3:3").Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp
        Range("F1").Select
        
        
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add(Range("F2:F78"), _
            xlSortOnFontColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(156, 0 _
            , 6)
        With ActiveWorkbook.Worksheets("Sheet1").Sort
            .SetRange Range("A1:N78")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        
        fTwo = Range("F2").value
        fThree = Range("F3").value
        
    Loop
    
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("F1:F83") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:N83")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    
End Sub
