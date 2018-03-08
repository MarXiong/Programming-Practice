Attribute VB_Name = "NoesysCallModule"
Public Sub NoesysCall(TextInput As String, ProductCodeInput As String, PackSizeInput As String, DateInput As String)
    
    Dim ResultsTable As ListObject
    Set ResultsTable = ActiveSheet.ListObjects("Results")
    
'    Dim CheckList As Object
'    Set CheckList = ActiveSheet.NoesysList
    
'    Clear out previous data in table
    If ResultsTable.ListRows.Count >= 1 Then
        ResultsTable.DataBodyRange.ClearContents
    End If
    
    Dim ResultsLocation As Range
    Set ResultsLocation = ResultsTable.Range
    
    Dim ListString As String
    
'    For i = 0 To CheckList.ListCount - 1
'        If CheckList.Selected(i) Then
'            ListString = ListString & ",[" & CheckList.List(i) & "] "
'        End If
'    Next
'    MsgBox ListString
    
    Dim StrSQL As String
    StrSQL = "SELECT TOP 20  (SELECT [Supplier].[Name] from [Noesys].[dbo].[Supplier] WHERE [Supplier].[ID]=[OrderRecords].[ID_Supplier]) as Supplier" & _
            " ,[ProductCode],[Description],[NamedPackSize],[PackPrice],[DateofPrice]," & _
            " (SELECT [ClientSize].[Name] from [Noesys].[dbo].[ClientSize] WHERE [ClientSize].[ID]=[OrderRecords].[ID_ClientSize]) as ClientSize " & _
            " FROM OrderRecords" & _
            " WHERE [Description] LIKE '%" & TextInput & "%'" & _
            " AND [DateofPrice]>= '" & DateInput & "'" & _
            " AND [ProductCode] LIKE '%" & ProductCodeInput & "%'" & _
            " AND [NamedPackSize] LIKE '%" & PackSizeInput & "%'" & _
            " AND [PackPrice]!=0" & _
            " ORDER BY [PackPrice];"


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
        ResultsLocation.Columns.AutoFit
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
        


End Sub


