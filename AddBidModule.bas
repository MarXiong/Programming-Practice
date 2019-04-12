Attribute VB_Name = "AddBidModule"
Sub AddNewBid()

Application.ScreenUpdating = False

Dim BidEvalTable As ListObject
Set BidEvalTable = Worksheets("Data").ListObjects("BidEvaluation")

Dim SupplierTable As ListObject
Set SupplierTable = Worksheets("Data").ListObjects("SupplierList")

Dim ActiveColumnRange As Range
Set ActiveColumnRange = BidEvalTable.HeaderRowRange.Find("Supplier Bids")

Dim SupplierChoiceTable As ListObject
Set SupplierChoiceTable = Worksheets("Data").ListObjects("SupplierChoiceSelection")

Dim GenericSupplierListString As String
GenericSupplierListString = "Add supplier name to supplier list"

Dim ColumnstoGroupRange As Range

Dim NumberofBidsInt As Integer
NumberofBidsInt = Application.CountIf(SupplierTable.ListColumns(1).DataBodyRange, "Bid*")

Dim BidColumnsInt As Integer
BidColumnsInt = Int(CountBidColumns(BidEvalTable) / NumberofBidsInt)

'Add as many columns as necessary for bid
Dim counter As Long
For counter = 1 To BidColumnsInt
    ActiveColumnRange.EntireColumn.Insert
Next counter

'Find new bid columns & group
Set ColumnstoGroupRange = ActiveColumnRange.Resize(, BidColumnsInt - 1).Offset(0, -BidColumnsInt)
ColumnstoGroupRange.Columns.Group

'Copy headers from BID1 and save to array
Dim BidColumnHeaderArray As Variant
Set BidColumnHeaderArray = BidEvalTable.HeaderRowRange.Find("Bid1").Resize(, BidColumnsInt)
BidColumnHeaderArray = BidColumnHeaderArray.Value

'Copy formulas over to new bid
Dim BPColumnFormulaArray As Variant
Set BPColumnFormulaArray = BidEvalTable.HeaderRowRange.Find("Bid1").Resize(, BidColumnsInt).EntireColumn
Application.Intersect(BPColumnFormulaArray, BidEvalTable.DataBodyRange).Copy

Intersect(ColumnstoGroupRange.EntireColumn, BidEvalTable.DataBodyRange).PasteSpecial (xlPasteAll)
On Error GoTo NoConstants
Intersect(ColumnstoGroupRange.EntireColumn, BidEvalTable.DataBodyRange).Cells.SpecialCells(xlCellTypeConstants).ClearContents

NoConstants:
    Resume Next

Application.Intersect(BPColumnFormulaArray, BidEvalTable.TotalsRowRange).Copy
Intersect(ColumnstoGroupRange.EntireColumn, BidEvalTable.TotalsRowRange).PasteSpecial (xlPasteAll)
On Error GoTo NoTotals
Intersect(ColumnstoGroupRange.EntireColumn, BidEvalTable.TotalsRowRange).Cells.SpecialCells(xlCellTypeConstants).ClearContents

NoTotals:
    Resume Next

'Work out number of new bid
Dim BidNameString As String
BidNameString = "BID" & NumberofBidsInt + 1

'Rename bid 1 headers and paste over new columns in table
For counter = 1 To BidColumnsInt
    BidColumnHeaderArray(1, counter) = Replace(BidColumnHeaderArray(1, counter), "BID1", BidNameString)
Next counter
Dim BidColumnHeaderRange As Range
Set BidColumnHeaderRange = Application.Intersect(ColumnstoGroupRange, BidEvalTable.HeaderRowRange)
BidColumnHeaderRange = BidColumnHeaderArray
BidColumnHeaderRange.Orientation = 0
BidColumnHeaderRange.VerticalAlignment = xlTop

Dim colour As Integer
colour = intRndColor()
BidColumnHeaderRange.Interior.ColorIndex = colour

ActiveColumnRange.Offset(0, -1).Value = BidNameString
ActiveColumnRange.Offset(0, -1).Interior.ColorIndex = colour

ColumnstoGroupRange.EntireColumn.AutoFit

'Add row to the supplier name list
Dim EndRowLong As Long
EndRowLong = SupplierTable.Range.End(xlDown).Row
Rows(EndRowLong + 1).Insert
Set NewRow = SupplierTable.ListRows.Add(NumberofBidsInt + 1, False)

SupplierTable.DataBodyRange(NumberofBidsInt + 1, 1).Value = BidNameString
SupplierTable.DataBodyRange(NumberofBidsInt + 1, 2).Value = GenericSupplierListString

SupplierTable.ListRows(EndRowLong).Range.Offset(1).Delete

'Extend Supplier Choice & Summary tables
Dim TableRange As Range
Set TableRange = SupplierChoiceTable.Range
SupplierChoiceTable.Resize TableRange.Resize(EndRowLong + 1)

Dim SummaryTable As ListObject
Set SummaryTable = Worksheets("Summary").ListObjects("BidSummary")

Set NewRow = SummaryTable.ListRows.Add(NumberofBidsInt + 1, False)
SummaryTable.DataBodyRange(NumberofBidsInt + 1, 1).Value = BidNameString

Application.ScreenUpdating = True

End Sub


Function CountBidColumns(Table As ListObject)

Dim FirstColumn As Integer
FirstColumn = Table.HeaderRowRange.Find("Bid1").Column

Dim LastColumn As Integer
LastColumn = Table.HeaderRowRange.Find("Supplier Bids").Column

CountBidColumns = LastColumn - FirstColumn

End Function

Function intRndColor()
   'USE - FUNCTION TO PICK RANDOM COLOR, ALSO ALLOWS EXCLUSION OF COLORS YOU DON'T LIKE
    Dim Again As Label
Again:
    intRndColor = Int((50 * Rnd) + 1) 'GENERATE RANDOM IN

    Select Case intRndColor
    Case Is = 1    'COLORS YOU DON'T WANT
        GoTo Again
    Case Is = pubPrevColor
        GoTo Again
    End Select
    pubPrevColor = intRndColor    'ASSIGN CURRENT COLOR TO PREV COLOR
End Function
