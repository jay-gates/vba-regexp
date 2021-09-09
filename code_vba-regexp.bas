' NOTE: THIS IS ACCESS VBA

Option Compare Database
Option Explicit

Public Function RegExp_GetFirstMatch( _
    ByVal patternText As String, _
    ByVal searchText As String, _
    Optional ByVal caseInsensitive As Boolean = True) As Variant
    ' Tests the regexp and returns the first match found, or Null
    ' SEE http://bytecomb.com/regular-expressions-in-vba/
    ' REQUIRES: Microsoft VBScript Regular Expressions 5.5 library reference
    ' Last edit 2014-12-28 by JG
    
    Dim re As New RegExp
    Dim matches As MatchCollection
    
    re.Global = False
    re.IgnoreCase = caseInsensitive
    re.pattern = patternText
    
    Set matches = re.Execute(searchText)
    
    If matches.Count = 0 Then ' none found
        RegExp_GetFirstMatch = Null
    Else
        RegExp_GetFirstMatch = matches(0).Value
    End If
    
    Set matches = Nothing
    Set re = Nothing
End Function

Function RegExp_GetFirstMatchGroup( _
    ByVal patternText As String, _
    ByVal searchText As String, _
    ByVal returnAllOnNomatch As Boolean, _
    Optional ByVal caseInsensitive As Boolean = True) As Variant
    ' Tests the regexp and returns the first match group, or Null
    ' Example: RegExp_GetFirstMatchGroup("(.*)(\Wtestb.*)","test testb testc", true) returns "test"
    ' REQUIRES: Microsoft VBScript Regular Expressions 5.5 library reference
    ' Last edit 2016-02-29 by JG
    
    Dim re As New RegExp
    Dim matches As MatchCollection
    
    re.Global = False
    re.IgnoreCase = caseInsensitive
    re.pattern = patternText
    
    Set matches = re.Execute(searchText)
    
    If matches.Count = 0 Then ' none found
        If returnAllOnNomatch Then ' return entire search text
            RegExp_GetFirstMatchGroup = searchText
        Else
            RegExp_GetFirstMatchGroup = Null
        End If
    Else ' return first match group
        RegExp_GetFirstMatchGroup = matches(0).SubMatches(0)
    End If
    
    Set matches = Nothing
    Set re = Nothing
End Function

Function RegExp_GetMatches(ByVal pattern As String, ByVal workString As Variant, Optional ByVal matchGlobal As Boolean = False) As Variant
    ' Returns Null for Null input, or Null if no match, or found match groups
    ' NOTE: Must specify capture groups
    ' REQUIRES: Microsoft VBScript Regular Expressions 5.5 library reference
    ' Last edit 2021-09-08 by JG
    
    Dim re As New RegExp
    Dim matchesCol As MatchCollection
    
    Dim i As Long
    Dim j As Long
    
    If IsNull(workString) Then
        RegExp_GetMatches = Null
    Else
        With re
           .Multiline = False
           .Global = matchGlobal
           .IgnoreCase = True
        End With
        
        re.pattern = pattern
        
        Set matchesCol = re.Execute(workString)
        
        RegExp_GetMatches = Null
        
        For i = 1 To matchesCol.Count
            For j = 1 To matchesCol(i - 1).SubMatches.Count
                RegExp_GetMatches = RegExp_GetMatches & matchesCol(i - 1).SubMatches(j - 1)
            Next j
        Next i
    End If
    
    Set re = Nothing
End Function

Function RegExp_Test(ByVal pattern As String, ByVal workString As Variant) As Variant
    ' Returns True if the work string matches the given regex, else False
    ' If workString is Null, returns Null
    ' REQUIRES: Microsoft VBScript Regular Expressions 5.5 library reference
    ' Last edit 2015-05-26 by JG
    
    Dim re As New RegExp
    
    If IsNull(workString) Then
        RegExp_Test = Null
    Else
        With re
           .Multiline = False
           .Global = False
           .IgnoreCase = True
           .pattern = pattern
        End With
        
        RegExp_Test = re.Test(workString)
    End If
    
    Set re = Nothing
End Function

Function RegExp_GetPhone_USA(ByVal workString As Variant) As Variant
    ' Returns Null for Null input, or Null if no match, or found number
    ' Matches (very "loose"):
    '   [...][(][ ]aaa[ ][)][ ][-][.][ ]bbb[ ][-][.][ ]cccc[...]
    ' REQUIRES: Microsoft VBScript Regular Expressions 5.5 library reference
    ' Last edit 2015-05-14 by JG
    
    Const pattern = "\(*\s*(\d{3})\s*\)*\s*\-*\.*\s*(\d{3})\s*\.*\-*\s*(\d{4})"
    
    Dim re As New RegExp
    Dim matchesCol As MatchCollection
    
    If IsNull(workString) Then
        RegExp_GetPhone_USA = Null
    Else
        With re
           .Multiline = False
           .Global = False
           .IgnoreCase = True
        End With
        
        re.pattern = pattern
        
        Set matchesCol = re.Execute(workString)
        
        If matchesCol.Count > 0 Then
'Debug.Print matchesCol(0).SubMatches(0)
'Debug.Print matchesCol(0).SubMatches(1)
'Debug.Print matchesCol(0).SubMatches(2)
        
            ' Format number uniformly
            RegExp_GetPhone_USA = _
                matchesCol(0).SubMatches(0) & "-" & _
                matchesCol(0).SubMatches(1) & "-" & _
                matchesCol(0).SubMatches(2)
        Else
            RegExp_GetPhone_USA = Null
        End If
    End If
    
    Set re = Nothing
End Function
