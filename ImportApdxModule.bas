Attribute VB_Name = "ImportApdxModule"
Sub SO()

Dim parentFolder As String

parentFolder = openDialog() & "\" '// change as required, keep trailing slash

Dim results As String

results = CreateObject("WScript.Shell").Exec("CMD /C DIR """ & parentFolder & "*.*"" /S /B /A:-D").StdOut.ReadAll

results = Replace(results, parentFolder, "")
'//Debug.Print results

'// uncomment to dump results into column A of spreadsheet instead:
'// Range("A1").Resize(UBound(Split(results, vbCrLf)), 1).Value = WorksheetFunction.Transpose(Split(results, vbCrLf))
'//-----------------------------------------------------------------
'// uncomment to filter certain files from results.
Dim filterType As String
filterType = "Grocery,xlsx"

Dim filters As Variant
filters = Split(filterType, ",")
Dim element As Variant

Dim filterResults As String

filterResults = results
For Each element In filters

    filterResults = Join(Filter(Split(filterResults, vbCrLf), element, True, vbTextCompare), vbCrLf)
Next element

Debug.Print filterResults
End Sub

Function openDialog()
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

   With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .filters.Clear
      '.Filters.Add "Excel", "*.xlsx"
      '.Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox

      End If
   End With
   openDialog = txtFileName
End Function


Function GetFilesIn(Folder As String, Optional Matching As String, Optional Unmatching As String) As Collection
  Dim FName As String
  Set GetFilesIn = New Collection
  If Matching = "" Then
    FName = Dir(Folder)
  Else
    FName = Dir(Folder & Matching)
  End If
  Do While FName <> ""
    If Unmatching = "" Then
      GetFilesIn.Add Folder & FName
    Else
      If Not FName Like Unmatching Then GetFilesIn.Add Folder & FName
    End If
    FName = Dir
  Loop
End Function


