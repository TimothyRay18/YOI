Attribute VB_Name = "ProcessButton"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub Process1()
    Dim wb As Workbook
    Dim dir1 As String
    dir1 = Range("B5").Value
    Set wb = Workbooks.Open(dir1)
    wb.Worksheets("OpenOrderFG").Activate

    Dim name1 As String
    name1 = ActiveWorkbook.Name

    ThisWorkbook.Activate
    Dim dir2 As String
    dir2 = Range("B9").Value
    Set wb = Workbooks.Open(dir2)
    wb.Worksheets("YOI").Activate

    Dim name2 As String
    name2 = ActiveWorkbook.Name

    Dim max_row As String

    Workbooks(name1).Activate
    max_row = getMaxRow(5)
    Range(Cells(4, 5), Cells(max_row, 5)).Select
    Selection.Copy
    Workbooks(name2).Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Workbooks(name1).Activate
    Range(Cells(4, 12), Cells(max_row, 12)).Select
    Selection.Copy
    Workbooks(name2).Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Workbooks(name1).Activate
    Range(Cells(4, 19), Cells(max_row, 20)).Select
    Selection.Copy
    Workbooks(name2).Activate
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("E2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("E2:E" + CStr(max_row - 2))
    
    Workbooks(name1).Activate
    Range(Cells(4, 14), Cells(max_row, 14)).Select
    Selection.Copy
    Workbooks(name2).Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("F:F").Select
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[-1])"
    Range("G2").Select
    Selection.NumberFormat = "General"
    Selection.AutoFill Destination:=Range("G2:G" + CStr(max_row - 2))
    Range("G2:G" + CStr(max_row - 2)).Select
    
    Workbooks(name1).Activate
    Range(Cells(4, 3), Cells(max_row, 3)).Select
    Selection.Copy
    Workbooks(name2).Activate
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveSheet.Range("$A$1:$H$" + CStr(max_row - 2)).AutoFilter Field:=6, Operator:= _
        xlFilterValues, Criteria1:="<" & CDbl(Date)
'
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Overdue"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Overdue"
    Range("F2:G2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    ActiveSheet.ShowAllData
    
'    Application.WorksheetFunction.WeekNum (Date)

    ActiveSheet.Range("$A$1:$H$" + CStr(max_row - 2)).AutoFilter Field:=6, Operator:= _
        xlFilterValues, Criteria1:="<>Overdue"
    Range("F1631").Select
    ActiveWorkbook.Worksheets("YOI").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("YOI").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "F2:F" + CStr(max_row - 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("YOI").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'    Get first visible row
    Dim visible_row As Double
    visible_row = Worksheets("YOI").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells().row
    Dim week1 As Double
    week1 = Range("G" + CStr(visible_row))
    Dim week2 As Double
    week2 = week1 + 11
    
'    Get Week of last day
    Dim weekNum As Double
    weekNum = CInt(Format(DateSerial(Year(dt), 12, 31), "ww", 2))
    
    If week2 > weekNum Then
        week2 = week2 - 53
        ActiveSheet.Range("$A$1:$H$" + CStr(max_row - 2)).AutoFilter Field:=7, Criteria1:="<" + CStr(week1), _
            Operator:=xlAnd, Criteria2:=">=" + CStr(week1)
    Else
        ActiveSheet.Range("$A$1:$H$" + CStr(max_row - 2)).AutoFilter Field:=7, Criteria1:="<" + CStr(week1), _
            Operator:=xlOr, Criteria2:=">=" + CStr(week2)
    End If
    
    visible_row = Worksheets("YOI").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells().row
    Range("G" + CStr(visible_row)).Select
    Selection.Copy
    Range(Cells(visible_row, 7), Cells(max_row - 2, 7)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ActiveSheet.Range("$A$1:$H$" + CStr(max_row - 2)).AutoFilter Field:=6, Operator:= _
        xlFilterValues, Criteria1:="Overdue"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Overdue"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Overdue"
    Range("F2:G2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    ActiveSheet.ShowAllData
        
    ActiveWorkbook.Worksheets("YOI").AutoFilter.Sort.SortFields.Clear
'    ActiveSheet.ShowAllData
    
    OldMacro.pivot
End Sub
Sub Process2()
    Dim nameYoi As String
    nameYoi = GetFilenameFromPath(Range("B9").Value)
    
    Dim dir1 As String
    dir1 = Range("B1").Value
    Workbooks.OpenText Filename:= _
        dir1, Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 2), _
        Array(3, 1), Array(4, 2), Array(5, 1), Array(6, 2), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1)), TrailingMinusNumbers:=True
        
    Rows("1:8").Select
    Selection.delete Shift:=xlUp
    Rows("2:2").Select
    Selection.delete Shift:=xlUp
    Columns("A:A").Select
    Selection.delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.delete Shift:=xlToLeft
    Columns("N:O").Select
    Selection.delete Shift:=xlToLeft
    Columns("P:Q").Select
    Selection.delete Shift:=xlToLeft
    Columns("M:M").Select
    Selection.delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.delete Shift:=xlToLeft
    Columns("N:N").Select
    Selection.delete Shift:=xlToLeft
    
    Range("A2").Select
    Range(Cells(2, 1), Cells(getMaxRow(1), 13)).Select
    Selection.Copy
    Workbooks(nameYoi).Activate
    Sheets("QTY PER").Select
    Range("B3").Select
    ActiveSheet.Paste
    
    Dim max_row As Double
    max_row = getMaxRow(2)
    Range("O3:Z3").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("O3:Z" + CStr(max_row))
    
    Range("C2").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$2:$Z$" + CStr(max_row)).AutoFilter Field:=11, Criteria1:="="
    ActiveSheet.Range("$A$2:$Z$" + CStr(max_row)).AutoFilter Field:=6, Criteria1:="<>ND"
    
    Range("B3").Select
    Range(Cells(3, 2), Cells(max_row, 2)).Select
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("A3").Select
    ActiveSheet.Paste
    
    Sheets("QTY PER").Select
    Range(Cells(3, 13), Cells(max_row, 13)).Select
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("B3").Select
    ActiveSheet.Paste
    
    Sheets("QTY PER").Select
    Range(Cells(3, 3), Cells(max_row, 3)).Select
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("C3").Select
    ActiveSheet.Paste
    
    Sheets("QTY PER").Select
    Range(Cells(3, 4), Cells(max_row, 4)).Select
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("D3").Select
    ActiveSheet.Paste
    
    Sheets("QTY PER").Select
    Range(Cells(3, 5), Cells(max_row, 5)).Select
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("E3").Select
    ActiveSheet.Paste
    
    Sheets("QTY PER").Select
    Range(Cells(3, 25), Cells(max_row, 25)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BOM by weekly").Select
    Range("F3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    OldMacro.BOM_SUM
    
    Range("G3:W3").Select
    Selection.AutoFill Destination:=Range("G3:W" + CStr(getMaxRow(1)))
    
    OldMacro.Last_Pivot_1
    
    Range("Q2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("Q2:Q" + CStr(getMaxRow(1)))
    Range("Q2:Q" + CStr(getMaxRow(1))).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
