Attribute VB_Name = "QABids"
Sub SearchColumns()

Dim EvalTable As ListObject
Set EvalTable = Worksheets("Product Pricing Data").ListObjects("ProductPricing")

Dim SupplierTable As ListObject
Set SupplierTable = Worksheets("Tender Summary").ListObjects("Summary")

Dim SupplierList As Collection
Set SupplierList = CollectUniques(SupplierTable.ListColumns("Bid No.").DataBodyRange)

SupplierList.Remove ("Mix1")
SupplierList.Remove ("Mix2")
SupplierList.Remove ("Mix3")


Dim BPAveragePriceColumn, BPDeviationColumn, BPDisregardColumn, BPNotesColumn, BPPackColumn As ListColumn

Set BPAveragePriceColumn = EvalTable.ListColumns("Average Wholesale Bid Price")
Set BPDeviationColumn = EvalTable.ListColumns("Standard Wholesale Bid Price Deviation")
Set BPDisregardColumn = EvalTable.ListColumns("Disregard this line?")
Set BPNotesColumn = EvalTable.ListColumns("Base Position Notes")
Set BPPackColumn = EvalTable.ListColumns("Pack Size")



Dim SearchName, CommentName, DisregardName, Factor As String

SearchName = "Difference %"
CommentName = "PP Query"
DisregardName = "Disregard "


Dim CurrentSearchColumn, CurrentCommentColumn, CurrentDisregardColumn, CurrentFactorColumn, CurrentPackColumn As ListColumn
Dim i As Integer


For Each Item In SupplierList
    On Error GoTo NextItem
        Set CurrentSearchColumn = EvalTable.ListColumns(Item & " " & SearchName)
        Set CurrentCommentColumn = EvalTable.ListColumns(Item & " " & CommentName)
        Set CurrentDisregardColumn = EvalTable.ListColumns(DisregardName & Item & "?")
        Set CurrentFactorColumn = EvalTable.ListColumns(Item & " Factor")
        Set CurrentPackColumn = EvalTable.ListColumns(Item & " Pack Size")
        
        If Not Application.WorksheetFunction.CountIf(CurrentSearchColumn.DataBodyRange, 0) = CurrentSearchColumn.DataBodyRange.Count Then
            For i = 1 To 10 'CurrentSearchColumn.DataBodyRange.Height
                If CurrentSearchColumn.DataBodyRange(i, 1) > BPAveragePriceColumn.DataBodyRange(i, 1) + 2 * BPDeviationColumn.DataBodyRange(i, 1) Or CurrentSearchColumn.DataBodyRange(i, 1) < (BPAveragePriceColumn.DataBodyRange(i, 1) - 2 * BPDeviationColumn.DataBodyRange(i, 1)) Then
                    CurrentDisregardColumn.DataBodyRange(i, 1).Value = "y"
                    CurrentCommentColumn.DataBodyRange(i, 1).Value = "Please confirm if this product is like for like?"
                ElseIf CurrentSearchColumn.DataBodyRange(i, 1) > 0.7 Then
                    CurrentDisregardColumn.DataBodyRange(i, 1).Value = "y"
                    
                    Factor = AddFactor.FactorPacks(BPPackColumn.DataBodyRange(i, 1), CurrentPackColumn.DataBodyRange(i, 1))
                    If IsNumeric(Factor) Then
                        CurrentCommentColumn.DataBodyRange(i, 1).Value = "Incorrect Factor"
                    End If
                    CurrentCommentColumn.DataBodyRange(i, 1).Value = "Please confirm this product has been priced correctly for the stated pack size"
                End If
            Next i
        End If
NextItem:
Next Item


End Sub

Public Function CollectUniques(rng As Range) As Collection
    
    Dim varArray As Variant, var As Variant
    Dim col As Collection
    
    'Guard clause - if Range is nothing, return a Nothing collection
    'Guard clause - if Range is empty, return a Nothing collection
    If rng Is Nothing Or WorksheetFunction.CountA(rng) = 0 Then
        Set CollectUniques = col
        Exit Function
    End If
        
    If rng.Count = 1 Then '<~ check for a single cell range
        Set col = New Collection
        col.Add Item:=CStr(rng.Value), Key:=CStr(rng.Value)
    Else '<~ otherwise the range contains multiple cells
        
        'Convert the passed-in range to a Variant array for SPEED and bind the Collection
        varArray = rng.Value
        Set col = New Collection
        
        'Ignore errors temporarily, as each attempt to add a repeat
        'entry to the collection will cause an error
        On Error Resume Next
        
            'Loop through everything in the variant array, adding
            'to the collection if it's not an empty string
            For Each var In varArray
                If CStr(var) <> vbNullString Then
                    col.Add Item:=CStr(var), Key:=CStr(var)
                End If
            Next var
    
        On Error GoTo 0
    End If
    
    'Return the contains-uniques-only collection
    Set CollectUniques = col
    
End Function


