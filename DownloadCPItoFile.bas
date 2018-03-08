Attribute VB_Name = "DownloadCPItoFile"
'Declaration of Windows API Function
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
 
Sub File_Download_From_Website()
    Dim myURL As String
    myURL = "https://www.ons.gov.uk/file?uri=/economy/inflationandpriceindices/datasets/consumerpriceinflation/current/consumerpriceinflationdetailedreferencetables.xls"
    Dim MyDate As String
    MyDate = ActiveWorkbook.Worksheets("New Index Input").Cells(2, 1).Text
    
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.Send
    
    myURL = WinHttpReq.ResponseBody
    If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.ResponseBody
    oStream.SaveToFile ("C:\Users\andrewm\Documents\FPI\" & MyDate & ".xls")
    oStream.Close
    End If
End Sub

