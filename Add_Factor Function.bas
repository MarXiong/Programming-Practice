Attribute VB_Name = "Add_Factor"
'Define collection to hold all acceptable combinations of units
    Public Collection As New Collection
    Public WeightUnitArray() As Variant
    Public VolumeUnitArray() As Variant
    Public SingleUnitArray() As Variant
    Public SingleUnitCheck As Boolean
    Public BPPackUnits As String
    Public BidPackUnits As String

Function Find_Equivalent_Prices(BPPack, BidPack)
    Dim Compatible As Boolean
    Dim BPPackValue As Variant
    Dim BidPackValue As Variant
    Dim BPPackUnits As String
    Dim BidPackUnits As String
    
    'Define array to hold valid combinations of units
    If Not Collection.Count > 0 Then
        WeightUnitArray = Array("g", "lbm", "kg", "oz")
        Collection.Add WeightUnitArray, "WeightUnitArray"
        VolumeUnitArray = Array("tsp", "oz", "pt", "gal", "l", "ml", "cl")
        Collection.Add VolumeUnitArray, "VolumeUnitArray"
        SingleUnitArray = Array("each", "single", "ea", "portion", "portions", "ptn", "sgl", "bunch", "case")
        Collection.Add SingleUnitArray, "SingleUnitArray"
    End If
    
    SingleUnitCheck = False
    
Continue:

    'Check whether pack sizes are the same
        'MsgBox BPPack
        'MsgBox BidPack
        
        If IsEmpty(BPPack) Or IsEmpty(BidPack) Then
            Find_Equivalent_Prices = "Pack size is not available, cannot compute factor"
        End If
        
        If BPPack = BidPack Then
            Find_Equivalent_Prices = 1
        End If
    
    'Read units for pack sizes & check they are valid units
        BPPackUnits = LCase(GetWords(BPPack))
        BidPackUnits = LCase(GetWords(BidPack))

        If Not BPPackUnits = "" Then
            BPPackUnits = CheckforUnitDescriptions(BPPackUnits)
            'MsgBox "BPPackUnits " & CStr(BPPackUnits(0))
        End If
        
        If Not BidPackUnits = "" Then
            BidPackUnits = CheckforUnitDescriptions(BidPackUnits)
            'MsgBox "BidPackUnits " & CStr(BidPackUnits(0))
        End If
        
        Compatible = CheckCompatibility(BPPackUnits, BidPackUnits)
        
    'Read string of numbers and x's from pack size
        BPPackValue = Trim(GetPackSize(Replace(BPPack, " ", "")))
        BidPackValue = Trim(GetPackSize(Replace(BidPack, " ", "")))
        
    'Check if multiple sets of numbers have been found in the pack size
        If InStr(BPPackValue, " ") Then
            TemporaryWordArray = SplitWords(BPPack)
            For j = LBound(TemporaryWordArray) To UBound(TemporaryWordArray)
                If TemporaryWordArray(j) = BPPackUnits Then
                    BPPackValue = Replace(TemporaryWordArray(j), BPPackUnits, "")
                End If
            Next j
            
        End If
        
        If InStr(BidPackValue, " ") Then
            TemporaryWordArray = SplitWords(BidPack)
            For j = LBound(TemporaryWordArray) To UBound(TemporaryWordArray)
                If TemporaryWordArray(j) = BidPackUnits Then
                    BidPackValue = Replace(TemporaryWordArray(j), BidPackUnits, "")
                End If
            Next j
        End If
        
        'MsgBox "BPPackValue " & CStr(BPPackValue(0))
        'MsgBox "BidPackValue " & CStr(BidPackValue(0))
        
    'Check whether pack sizes need converting to valid numbers
        If Not IsNumeric(BPPackValue) Then
            If Not IsError(Application.Evaluate(Replace(LCase(BPPackValue), "x", "*"))) Then
                BPPackValue = Application.Evaluate(Replace(LCase(BPPackValue), "x", "*"))
            ElseIf InStr(1, BPPackValue, "-") Then
                BPPackValue = Application.Evaluate(Replace(LCase(ReadHyphens(BPPackValue)), "x", "*"))
            End If
        End If
        If Not IsNumeric(BidPackValue) Then
            If Not IsError(Application.Evaluate(Replace(LCase(BidPackValue), "x", "*"))) Then
                BidPackValue = Application.Evaluate(Replace(LCase(BidPackValue), "x", "*"))
            ElseIf InStr(1, BidPackValue, "-") Then
                BidPackValue = Application.Evaluate(Replace(LCase(ReadHyphens(BidPackValue)), "x", "*"))
            End If
        End If
        'MsgBox "BPPackValue " & CStr(BPPackValue)
        'MsgBox "BidPackValue " & CStr(BidPackValue)
        
    'Check whether the units of the pack sizes match and if not convert the bid pack size value
        If BPPackValue = "" And BPPackUnits = "" Then
            Find_Equivalent_Prices = "Cannot factor as pack size is missing"
        ElseIf BidPackValue = "" And BidPackUnits = "" Then
            Find_Equivalent_Prices = ""
        ElseIf Compatible Then
            If Not SingleUnitCheck Then
                If BPPackValue = "" Then
                    BPPackValue = 1
                End If
                If BidPackValue = "" Then
                    BidPackValue = 1
                End If
                If StrComp(WorksheetFunction.Clean(BPPackUnits), WorksheetFunction.Clean(BidPackUnits), vbTextCompare) = 0 Then
                    Find_Equivalent_Prices = " =" & BPPackValue & "/" & BidPackValue
                End If
                On Error Resume Next
                    Find_Equivalent_Prices = " =" & BPPackValue & "/" & WorksheetFunction.Convert(BidPackValue, BidPackUnits, BPPackUnits)
            Else:
                If BPPackValue = "" Then
                    BPPackValue = 1
                End If
                If BidPackValue = "" Then
                    BidPackValue = 1
                End If
                    Find_Equivalent_Prices = " =" & BPPackValue & "/" & BidPackValue
            End If
        Else:
        If BPPackUnits = "" And BidPackUnits = "" Then
            Find_Equivalent_Prices = BPPackValue / BidPackValue
        End If
            Find_Equivalent_Prices = "Cannot factor as pack sizes are incompatible"
        End If
    
Exit Function

End Function


Function ReadHyphens(ByVal strIn As String) As Variant  'Convert hyphens into average value
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .Pattern = "([0-9]*(\-[0-9]+)+)"
        NumStr = .Replace(strIn, "($1)/2")
    End With
    
    NumStr = Replace(NumStr, "-", "+")
    ReadHyphens = Trim(NumStr)
End Function

Function GetPackSize(ByVal strIn As String) As String  'Array of numeric strings split by x
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^xX0-9\-\.]+"
        NumStr = .Replace(strIn, " ")
    End With
    With RegExpObj
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b[xX]+|[xX]+\b"
        NumStr = .Replace(NumStr, "")
    End With
    
    If NumStr = " *" Then
        NumStr = "1"
    End If
    
    GetPackSize = NumStr
End Function

Function SplitWords(ByVal strIn As String) As Variant  'Return words in string
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .Pattern = "[^\w\.\-]+"
        NumStr = .Replace(strIn, " ")
    End With
        
    GetNums = Trim(NumStr)
    SplitWords = Split(NumStr, " ")
End Function

Function GetWords(ByVal strIn As String) As Variant  'Array of word strings
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .Pattern = "[xX0-9\-\.]+"
        NumStr = .Replace(strIn, " ")
    End With
    With RegExpObj
        .Global = True
        .Pattern = "[^\w]+"
        NumStr = .Replace(NumStr, " ")
    End With
        
    NumStr = Replace(NumStr, "*x*", " ")
    GetWords = Trim(NumStr)

End Function

Function CheckforUnitDescriptions(ByVal UnitsInput As String) As String 'Check for units and convert them to correct format
    Dim UnitReplacement As String
    UnitReplacement = Replace(UnitsInput, "ltr", "l")
    UnitReplacement = Replace(UnitReplacement, "gm", "g")
    UnitReplacement = Replace(UnitReplacement, "kilo", "kg")
    UnitReplacement = Replace(UnitReplacement, "kgm", "kg")
    CheckforUnitDescriptions = UnitReplacement

End Function

Function CheckCompatibility(BPUnits As String, BidUnits As String) As Boolean 'Check for whether units are compatible
    Dim Item As Variant
    SingleUnitCheck = False
    
    If BPUnits = BidUnits Then
        CheckCompatibility = True
        If IsInArray(BPUnits, SingleUnitArray) Or BPUnits = "" Then
            SingleUnitCheck = True
        End If
        Exit Function
    End If
    
    If BPUnits = "" And IsInArray(BidUnits, SingleUnitArray) Then
        CheckCompatibility = True
        SingleUnitCheck = True
        Exit Function
    End If
    
    If BidUnits = "" And IsInArray(BPUnits, SingleUnitArray) Then
        CheckCompatibility = True
        SingleUnitCheck = True
        Exit Function
    End If
    
    For Each Item In Collection
        For i = LBound(Item) To UBound(Item)
            For j = LBound(Item) To UBound(Item)
                If BPUnits = Item(i) And BidUnits = Item(j) Then
                    CheckCompatibility = True
                    BPUnits = Item(i)
                    BidUnits = Item(j)
                    If IsInArray(CStr(Item(j)), SingleUnitArray) Then
                        SingleUnitCheck = True
                    End If
                    Exit Function
                End If
            Next j
        Next i
    Next Item
    
'    For Each Item In Collection
'        If IsInArray(BPUnits, Item) And IsInArray(BidUnits, Item) Then
'        CheckCompatibility = True
'            If IsInArray(CStr(Item(BPUnits)), SingleUnitArray) Then
'                SingleUnitCheck = True
'            End If
'            Exit Function
'        End If
'    Next Item
    
    CheckCompatibility = False
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean 'Check whether an element is in an array
  IsInArray = IsNumeric(Application.Match(stringToBeFound, arr, 0))
End Function
