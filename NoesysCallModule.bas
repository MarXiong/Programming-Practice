Attribute VB_Name = "NoesysCallModule"
Public Sub NoesysCall(DescriptionInput As String, ProductCodeInput As String, PackSizeInput As String, DateInput As String, SupplierInput As String, SourceInput As String)
    
    Application.ScreenUpdating = False
    Dim ResultsTable As ListObject
    Set ResultsTable = ActiveSheet.ListObjects("Results")
    Dim TableRows As Integer
    TableRows = ResultsTable.ListRows.Count
    
'    Dim CheckList As Object
'    Set CheckList = ActiveSheet.NoesysList
    
'    Clear out previous data in table
    If TableRows >= 1 Then
        ResultsTable.DataBodyRange.ClearContents
    End If
    
    Dim ResultsLocation As Range
    Set ResultsLocation = ResultsTable.Range.Offset(1, 0)
    
    Dim ListString As String
    
'    For i = 0 To CheckList.ListCount - 1
'        If CheckList.Selected(i) Then
'            ListString = ListString & ",[" & CheckList.List(i) & "] "
'        End If
'    Next
'    MsgBox ListString
    
    
    DescriptionInput = SplitInputStrings.SplitInputString(DescriptionInput, "[Description]")
    If Not ProductCodeInput = "" Then
        ProductCodeInput = " AND " & SplitInputStrings.SplitInputString(ProductCodeInput, "[ProductCode]")
    End If
    If Not PackSizeInput = "" Then
        PackSizeInput = " AND " & SplitInputStrings.SplitInputString(PackSizeInput, "[NamedPackSize]")
    End If
    If Not SupplierInput = "" Then
        SupplierInput = " AND " & SplitInputStrings.SplitInputString(SupplierInput, "[Supplier].[Name]")
    End If
    If Not SourceInput = "" Then
        SourceInput = " AND " & SplitInputStrings.SplitInputString(SourceInput, "[Category7]")
    End If
    
    'Construct the SQL string to pass to Noesys
    Dim StrSQL As String
    StrSQL = "SELECT TOP " & TableRows & " [Supplier].[Name] as Supplier" & _
            " ,[Category7] as Source ,[ProductCode] as 'Product Code',[Description],[NamedPackSize] as 'Pack Size',[PackPrice] as Price,[DateofPrice] as Date " & _
            " FROM OrderRecords,Supplier" & _
            " WHERE " & DescriptionInput & _
            " AND [DateofPrice]>= '" & DateInput & "'" & _
            ProductCodeInput & _
            PackSizeInput & _
            SupplierInput & _
            SourceInput & _
            " AND [PackPrice]!=0" & _
            " AND [Supplier].[ID]=[OrderRecords].[ID_Supplier]" & _
            " ORDER BY [PackPrice];"
    'MsgBox StrSQL

    ' Create a recordset object.
    If cnPubs Is Nothing Then
        Call Connection
    End If
    
    Dim rsPubs As ADODB.Recordset
    Set rsPubs = New ADODB.Recordset
    Set rsPubs = cnPubs.Execute(StrSQL)

    ' Check we have data.
    If Not rsPubs.EOF Then
    ' Transfer result.
        ResultsLocation.CopyFromRecordset rsPubs
        'ResultsLocation.Columns.AutoFit
'        Add in column names from DB
        Dim index As Integer
        For index = 0 To rsPubs.Fields.Count - 1
            ResultsTable.HeaderRowRange(1, index + 1).Value = rsPubs.Fields(index).Name
        Next
        ' Close the recordset
        rsPubs.Close
            
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If

'    Dim RecordList As ADODB.Recordset
'    Set RecordList = cnPubs.Execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS")
'    With CheckList
'        .Clear
'        RecordList.MoveFirst
'        Do Until RecordList.EOF
'                .AddItem RecordList!column_name
'                RecordList.MoveNext
'        Loop
'    End With
'
'    With CheckList
'    If .ListCount = 0 Then GoTo Here
'        RecordList.MoveFirst
'        Do Until RecordList.EOF
'            For j = 0 To CheckList.ListCount - 1
'                If .Selected(j) = RecordList!column_name Then
'                    RecordList.MoveNext
'                Else
'Here:
'                    .AddItem RecordList!column_name
'                    RecordList.MoveNext
'                End If
'            Next
'        Loop
'    End With

    Application.ScreenUpdating = True
Exit Sub

End Sub


