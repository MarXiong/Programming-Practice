Attribute VB_Name = "Module5"
Sub RunRscript()
'Runs external R code through Shell
'The location of the script is 'C:\R'
'Script name is 'Hello.R'

Dim shell As Object
Set shell = VBA.CreateObject("WScript.Shell")

Dim waitTillComplete As Boolean: waitTillComplete = True
Dim style As Integer: style = 1
Dim errorCode As Integer

Dim var1, var2 As Double
var1 = Worksheets("Noesys").Range("B2").Value
var2 = Worksheets("Noesys").Range("C2").Value

Dim path As String
path = "RScript C:/Users/andrewm/Documents/R/Test/Hello.R " & var1 & " " & var2

errorCode = shell.Run(path, style, waitTillComplete)

End Sub
