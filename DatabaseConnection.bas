Attribute VB_Name = "DatabaseConnection"
Public cnPubs As ADODB.Connection
    
Public Sub Connection()
    
    Dim ReturnArray
    
    ' Create a connection object.
    Set cnPubs = New ADODB.Connection
    
    ' Provide the connection string.
    Dim strConn As String
    
    'Use the SQL Server OLE DB Provider.
    strConn = "PROVIDER=SQLOLEDB.1;"
    
    'Connect to the Pubs database on the local server.
'    strConn = strConn & "DATA SOURCE=PRESTIGELAP001\MYDATABASE;INITIAL CATALOG=Noesys;"

    'Connect to shared Noesys server.
    strConn = strConn & "DATA SOURCE=PRESTIGE-SBS\SQLEXPRESS;INITIAL CATALOG=Noesys;"
    
    'Use an integrated login.
    strConn = strConn & "INTEGRATED SECURITY=sspi;"
    
    'Now open the connection.
    cnPubs.Open strConn
    'See if it worked
    If cnPubs.State = adStateOpen Then
        MsgBox "Successfully Connected!"
    Else
        MsgBox "Sorry. No Connection Established."
    End If
            
End Sub
