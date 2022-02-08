Attribute VB_Name = "validationHelper"
Public Function birthdayExtract(birthday As String) As String
    
    Dim re, matches, oMatch
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "([\d]{2})[\/-]?([\d]{2})[\/-]?([\d]{2,4})"
    re.Global = True
    Dim formatString As String
    formatString = ""
    Set matches = re.Execute(birthday)
    
    If matches.count = 0 Then
        Err.Raise Number:=vbObjectError + 513, Description:="Wrong Birthday"
    End If
    
    Set oMatch = matches(0)
    
    
    
    
    formatString = oMatch.SubMatches(0) & "-" & oMatch.SubMatches(1) & "-" & oMatch.SubMatches(2)
    
    birthdayExtract = formatString
    
    Set re = Nothing
    
    
    
End Function
