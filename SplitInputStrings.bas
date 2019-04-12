Attribute VB_Name = "SplitInputStrings"
'Split text keywords into include and exclude keywords
Function SplitInputString(TextInput As String, DatabaseColumn As String)
    Dim SplitText() As String
    Dim SplitLikeText As String
    Dim SplitNotLikeText As String
    SplitLikeText = ""
    SplitNotLikeText = ""
    
    If InStr(TextInput, ",") Then
        SplitText() = Split(TextInput, ",")
        If UBound(SplitText) = LBound(SplitText) Then
            If InStr(SplitText(1), "-") Then
                SplitNotLikeText = SplitNotLikeText & DatabaseColumn & " NOT LIKE " & "'" & Replace(SplitText(1), "-", "%") & "%' "
                TextInput = SplitNotLikeText
            Else
                SplitLikeText = SplitLikeText & DatabaseColumn & " LIKE " & "'%" & SplitText(1) & "%' "
                TextInput = SplitLikeText
            End If
        Else
            For i = LBound(SplitText) To UBound(SplitText)
                If InStr(SplitText(i), "-") Then
                    SplitNotLikeText = SplitNotLikeText & DatabaseColumn & " NOT LIKE " & "'" & Replace(SplitText(i), "-", "%") & "%' "
                Else
                    SplitLikeText = SplitLikeText & DatabaseColumn & " LIKE " & "'%" & SplitText(i) & "%' "
                End If
            Next i
            SplitLikeText = Replace(SplitLikeText, " " & DatabaseColumn, " AND " & DatabaseColumn)
            SplitNotLikeText = Replace(SplitNotLikeText, " " & DatabaseColumn, " AND " & DatabaseColumn)
            If SplitNotLikeText = "" Then
                TextInput = SplitLikeText
            ElseIf SplitLikeText = "" Then
                TextInput = SplitNotLikeText
            Else
                TextInput = SplitLikeText & " AND " & SplitNotLikeText
            End If
        End If
        MsgBox TextInput
        SplitInputString = TextInput
    Else
        If InStr(TextInput, "-") Then
            TextInput = DatabaseColumn & " NOT LIKE " & "'%" & Replace(TextInput, "-", "") & "%'"
        Else
            TextInput = DatabaseColumn & " LIKE " & "'%" & TextInput & "%'"
        End If
            SplitInputString = TextInput
    End If

End Function
