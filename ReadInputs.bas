Attribute VB_Name = "ReadInputs"
Public Sub ButtonClick()
    Dim DescriptionInput As String
    Dim DateLimit As String
    Dim PackInput As String
    Dim ProductCodeInput As String
    
    DescriptionInput = ActiveSheet.DescriptionBox.Text
    PackInput = ActiveSheet.PackSizeBox.Text
    ProductCodeInput = ActiveSheet.ProductCodeBox.Text
    DateLimit = ActiveSheet.DateBox.Value
    
    'Format input as date
    If IsDate(DateLimit) Then
        DateLimit = Format(DateLimit, "mm/dd/yy")
        Call NoesysCallModule.NoesysCall(DescriptionInput, ProductCodeInput, PackInput, DateLimit)
    Else: On Error GoTo GoHere
        DateLimit = Format(DateLimit, "dd/mm/yy")
        DateLimit = Format(DateLimit, "dd/mm/yyyy")
        DateLimit = Format(DateLimit, "mm/dd/yy")
        DateLimit = Format(DateLimit, "mm/dd/yyyy")
        DateLimit = Format(DateLimit, "mm/dd/yy")
        Call Module1.NoesysCall(DescriptionInput, ProductCodeInput, PackInput, DateLimit)
        
GoHere:
        MsgBox "Invalid Date Provided"
    End If
    

        
End Sub
