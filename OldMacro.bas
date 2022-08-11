Attribute VB_Name = "OldMacro"
Sub BOM_SUM()
'
' BOM_SUM Macro
'

'
    Range("G3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-6],YOI!C10:C30,2,0))*RC[-1]),0)"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-7],YOI!C10:C30,3,0))*RC[-2]),0)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-8],YOI!C10:C30,4,0))*RC[-3]),0)"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-9],YOI!C10:C30,5,0))*RC[-4]),0)"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-10],YOI!C10:C30,6,0))*RC[-5]),0)"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-11],YOI!C10:C30,7,0))*RC[-6]),0)"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-12],YOI!C10:C30,8,0))*RC[-7]),0)"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-13],YOI!C10:C30,9,0))*RC[-8]),0)"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-14],YOI!C10:C30,10,0))*RC[-9]),0)"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-15],YOI!C10:C30,11,0))*RC[-10]),0)"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-16],YOI!C10:C30,12,0))*RC[-11]),0)"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-17],YOI!C10:C30,13,0))*RC[-12]),0)"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(((VLOOKUP(RC[-18],YOI!C10:C30,14,0))*RC[-13]),0)"
    Range("T3").Select
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-20],'List FG'!C[-20],1,0)"
    Range("U3").Select
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=YOI!R[1]C[4]"
    Range("G1").Select
    Selection.Copy
    Range("G1:S1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("H1:L1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("H1:L1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-13]:RC[-1])"
    Range("T4").Select
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-21],YOI!C[-21]:C[-13],8,0)"
    Range("V3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-21],YOI!C[-21]:C[-14],8,0)"
    Range("W3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-22],YOI!C[-22]:C[-14],9,0)"
    Range("W3").Select
    Range("S1").Select
    Selection.Copy
    Range("H1:S1").Select
    Range("S1").Activate
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
End Sub
Sub Last_Pivot_1()
'
' Last_Pivot_1 Macro
'

'
    Columns("G:U").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("T7").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("BOM by weekly").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BOM by weekly").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("U2:U42351"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("BOM by weekly").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("U3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Columns("U:U").Select
    Selection.Replace What:="#n/a", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("H38471").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    ActiveWindow.SmallScroll Down:=-15
    Columns("C:S").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BOM by weekly!R2C3:R1048576C19", Version:=6).CreatePivotTable _
        TableDestination:="Pivot!R1C1", TableName:="PivotTable4", DefaultVersion _
        :=6
    Sheets("Pivot").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable4")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    ActiveWindow.FreezePanes = False
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Component number")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Follow-up mat")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Obj.Desc")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Overdue"), "Sum of Overdue", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W"), "Sum of W", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W1"), "Sum of W1", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W2"), "Sum of W2", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W3"), "Sum of W3", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W4"), "Sum of W4", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W5"), "Sum of W5", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W6"), "Sum of W6", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W7"), "Sum of W7", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W8"), "Sum of W8", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W9"), "Sum of W9", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W10"), "Sum of W10", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W11"), "Sum of W11", xlSum
    Sheets("Pivot").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Component number"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Component number"). _
        LayoutForm = xlTabular
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Follow-up mat").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Follow-up mat").LayoutForm _
        = xlTabular
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Cells.Select
    Selection.Copy
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Sheets("BOM by weekly").Select
    Range("G1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Pivot").Select
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("Q1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "SUM"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "SUPPLIER"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "PIC"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "FG"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-13]:RC[-1])"
    Range("S2").Select
    
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-19],'BOM by weekly'!C[-17]:C[1],19,0)"
    Range("T3").Select
    Sheets("BOM by weekly").Select
    Columns("V:V").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Pivot").Select
    Range("R9").Select
    Columns("B:B").Select
    Cells.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("J19").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    ActiveWindow.SmallScroll Down:=-18
    Columns("B:B").Select
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.End(xlUp).Select
    Range("A2").Select
    Selection.End(xlDown).Select
    Range("B5509").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("S2").Select
    
    

End Sub
Sub pivot()
'
' pivot Macro
'

'
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("A:G").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "YOI!R1C1:R1048576C7", Version:=6).CreatePivotTable TableDestination:= _
        "YOI!R1C10", TableName:="PivotTable10", DefaultVersion:=6
    Sheets("YOI").Select
    Cells(1, 10).Select
    With ActiveSheet.PivotTables("PivotTable10")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable10").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable10").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable10").PivotFields("Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable10").AddDataField ActiveSheet.PivotTables( _
        "PivotTable10").PivotFields("Qty"), "Count of Qty", xlCount
    With ActiveSheet.PivotTables("PivotTable10").PivotFields("Week")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable10").PivotFields("Count of Qty")
        .Caption = "Sum of Qty"
        .Function = xlSum
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Range("Y2").Select
    ActiveSheet.PivotTables("PivotTable10").PivotFields("Week").PivotItems( _
        "overdue").Position = 1
    Range("K4").Select
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1

 
    Range("K3").Select
End Sub
Sub delete()
'
' delete Macro
'

'
    Cells.Select
    Selection.ClearContents
    Sheets("BOM by weekly").Select
    Rows("4:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("YOI").Select
    ActiveWindow.ScrollRow = 1
    Columns("J:AB").Select
    Selection.ClearContents
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("YOI").Select
End Sub


