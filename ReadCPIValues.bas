Attribute VB_Name = "ReadCPIValues"
Sub Copy_CPI_Values()
    Application.ScreenUpdating = False
    Dim path As String
    Dim openWb As Workbook
    Dim openWs As Worksheet
    Dim MyDate As String
    Dim PasteArea As Range
    
    MyDate = ActiveWorkbook.Worksheets("New Index Input").Cells(2, 1).Text
    Set PasteArea = ActiveWorkbook.Worksheets("New Index Input").Cells(5, 10)

    path = "C:\Users\andrewm\Documents\FPI\" & MyDate & ".xls"

    Set openWb = Workbooks.Open(path)
    Set openWs = openWb.Sheets("Table 4")

    With openWs
        .Cells(11, 11).Resize(13, 1).Copy
        PasteArea.PasteSpecial
        .Cells(30, 11).Resize(19, 1).Copy
        PasteArea.Offset(15, 0).PasteSpecial
    End With

    Dim rng As Range

    

    PasteArea.Resize(rng.Rows.Count) = rng
    
    With openWs
        .Cells(11, 6).Resize(13, 1).C5opy
        PasteArea.Offset(0, -1).PasteSpecial
        .Cells(30, 6).Resize(19, 1).Copy
        PasteArea.Offset(15, -1).PasteSpecial
    End With

'Store blank cells inside a variable
      On Error GoTo NoBlanksFound
        Set rng = PasteArea.EntireColumn.SpecialCells(xlCellTypeBlanks)
      On Error GoTo 0
    
    'Delete blank cells and shift upward
      rng.Rows.Delete Shift:=xlShiftUp
    
    'ERROR HANLDER
NoBlanksFound:
      MsgBox "No Blank cells were found"

    openWb.Close (True)
    Application.ScreenUpdating = True
End Sub
