Attribute VB_Name = "Update_Index"
Sub CopyRange()

Dim source As Worksheet
Dim destination As Worksheet
Dim datasheet As Worksheet
Dim ConvertSheetemptyColumn As Long
Dim DataSheetemptyColumn As Long

'Find each worksheet we are going to use
Set source = ActiveWorkbook.Worksheets("New Index Input")
Set destination = ActiveWorkbook.Sheets("Index Old to New Values")
Set datasheet = ActiveWorkbook.Sheets("Data")

'Find first empty column in destination (actually cell in Row 1)'
ConvertSheetemptyColumn = destination.Cells(1, destination.Columns.Count).End(xlToLeft).Column
If ConvertSheetemptyColumn > 1 Then
ConvertSheetemptyColumn = ConvertSheetemptyColumn + 1
End If

'Fill across from the last column in destination to the next
destination.Cells(1, ConvertSheetemptyColumn - 1).EntireColumn.AutoFill destination:=destination.Cells(1, ConvertSheetemptyColumn - 1).EntireColumn.Resize(, 2)

'Copy across the index numbers given in Source
source.Range("C5:C18").Copy destination.Cells(2, ConvertSheetemptyColumn)

'Look at A3 in source to find column in Data referring to the new month
DataSheetemptyColumn = source.Cells(3, 1).Value

'Paste ongoing FPI index values into the Data worksheet, in the column given by DataSheetemptyColumn
destination.Cells(30, ConvertSheetemptyColumn).Resize(14).Copy
datasheet.Cells(33, DataSheetemptyColumn).PasteSpecial xlPasteValues

MsgBox "Index has been updated succesfully!" & vbNewLine & "This month is " & source.Cells(2, 1).Text

End Sub
