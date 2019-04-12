Attribute VB_Name = "FillReports"
Sub FillReports()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Dim Path As String
Dim wb As Workbook

Dim DestSheet, TableRange As String
Dim Column As ListColumn

Dim CopyTable, DestTable As ListObject

Dim FilterColumnNum As Integer

Dim Filelist As Variant

Dim CurrentWorkbook As String

Dim SheetList, Match As Variant


Path = ActiveWorkbook.Path

CurrentWorkbook = ActiveWorkbook.Name

Path = Path & "\"
Dim Filename As String
Filename = Dir(Path & "*Product Line Detail*.xl??", vbNormal)

SheetList = Split("Included,Excluded", ",")
Match = Split("Yes,No", ",")

Dim i, TotalRows As Integer

Do While Filename <> ""
    For i = 0 To UBound(SheetList)
        Debug.Print (i)
        Debug.Print (SheetList(i))
        Debug.Print (Match(i))
        Set wb = Application.Workbooks.Open(Path & Filename)
        Set DestTable = wb.Sheets(SheetList(i)).ListObjects(1)
        Debug.Print (DestTable.Name)
        
        Set CopyTable = Workbooks(CurrentWorkbook).Sheets("Matchmaker").ListObjects(1)
        
        FilterColumnNum = Application.WorksheetFunction.Match("Match?", CopyTable.HeaderRowRange, 0)
        TotalRows = Application.WorksheetFunction.CountIf(CopyTable.ListColumns("Match?").DataBodyRange, Match(i))
        Debug.Print (TotalRows)
        
        If DestTable.DataBodyRange.rows.Count < TotalRows Then
            TableRange = DestTable.Range.Address
            TableRange = Left(TableRange, FindN("$", TableRange, 4))
            TableRange = TableRange & Trim(Str(TotalRows))
            DestTable.Resize Range(TableRange)
        End If
        
        CopyTable.Parent.AutoFilterMode = False
        CopyTable.Range.AutoFilter Field:=FilterColumnNum, Criteria1:=Match(i), VisibleDropDown:=False
        
        For Each Column In DestTable.ListColumns
            CopyTable.ListColumns(Column.Name).DataBodyRange.SpecialCells(xlVisible).Copy
            Column.DataBodyRange.PasteSpecial (xlPasteValues)
        Next
        
        Dim EmptyTableRange As Range
        
        On Error Resume Next
        Set EmptyTableRange = DestTable.ListColumns(1).Range.SpecialCells(xlCellTypeBlanks)
        If Not EmptyTableRange Is Nothing Then
            EmptyTableRange.Delete Shift:=xlUp
        End If
        
    Next
    
    wb.Save
    wb.Close
    DoEvents
    Filename = Dir()
    
Loop
Workbooks(CurrentWorkbook).RefreshAll
Debug.Print

Application.ScreenUpdating = True
Application.DisplayAlerts = True
EndSub:
End Sub





