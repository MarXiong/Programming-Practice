Attribute VB_Name = "LevenshteinFuzzyMatch"
' Levenshtein tweaked for UTLIMATE speed and CORRECT results
' Solution based on Longs
' Intermediate arrays holding Asc()make difference
' even Fixed length Arrays have impact on speed (small indeed)
' Levenshtein version 3 will return correct percentage


Public Function ClosestString(Lookup As String, SearchArray As Range)
Dim Ranking As Variant
Ranking = SearchArray
ReDim Preserve Ranking(1 To SearchArray.Count, 1 To 2)

Dim MaxVal As Double
MaxVal = 0
Dim MaxIndex As Integer

searchlength = SearchArray.Count

For i = 1 To searchlength
    Ranking(i, 2) = Levenshtein(Lookup, CStr(Ranking(i, 1)))
    If MaxVal < Ranking(i, 2) Then
        MaxVal = Ranking(i, 2)
        MaxIndex = i
    End If
Next

Dim Answer As String
Answer = CStr(Ranking(MaxIndex, 1))

ClosestString = Answer
 
End Function


Public Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

Dim i As Long, j As Long, string1_length As Long, string2_length As Long
Dim distance(0 To 100, 0 To 100) As Long, smStr1(1 To 100) As Long, smStr2(1 To 100) As Long
Dim min1 As Long, min2 As Long, min3 As Long, minmin As Long, MaxL As Long

string1_length = Len(string1):  string2_length = Len(string2)

distance(0, 0) = 0

For i = 1 To string1_length:    distance(i, 0) = i: smStr1(i) = Asc(LCase(Mid$(string1, i, 1))): Next
For j = 1 To string2_length:    distance(0, j) = j: smStr2(j) = Asc(LCase(Mid$(string2, j, 1))): Next
For i = 1 To string1_length
    For j = 1 To string2_length
        If smStr1(i) = smStr2(j) Then
            distance(i, j) = distance(i - 1, j - 1)
        Else
            min1 = distance(i - 1, j) + 1
            min2 = distance(i, j - 1) + 1
            min3 = distance(i - 1, j - 1) + 1
            If min2 < min1 Then
                If min2 < min3 Then minmin = min2 Else minmin = min3
            Else
                If min1 < min3 Then minmin = min1 Else minmin = min3
            End If
            distance(i, j) = minmin
        End If
    Next
Next

' Levenshtein will properly return a percent match (100%=exact) based on similarities and Lengths etc...
MaxL = string1_length: If string2_length > MaxL Then MaxL = string2_length
Levenshtein = 100 - CLng((distance(string1_length, string2_length) * 100) / MaxL)

End Function
