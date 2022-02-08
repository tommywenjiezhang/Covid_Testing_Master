Attribute VB_Name = "generateIndex"
Public Function generateIndex(ByVal c As String) As Long
    Dim o As New Dictionary
    Dim rng As Range
    Dim last_row As Long
    Dim out As Long
    With empBirthday
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        Set rng = .Range("A2:A" & last_row)
    End With
    
    Dim letter As String
    Dim idx As Long
    For Each cell In rng
        letter = Left(cell.value, 1)
        If Not o.Exists(UCase(letter)) Then
            o(letter) = 1
        Else
            o(letter) = o(letter) + 1
        End If
    Next cell
    out = o(UCase(c))
    out = CLng(checkIndexExisted(c, out))
    
    generateIndex = out
End Function

Private Function checkIndexExisted(empInitial As String, empNumber As Long) As Long
    Dim lookup_value As Variant
    Dim last_row As Long
    Dim rng As Range
    
    With empBirthday
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        Set rng = .Range("A1:A" & last_row)
    End With
    lookup_value = Application.VLookup((empInitial + CStr(empNumber)), rng, 1, False)
    If IsError(lookup_value) Then
        checkIndexExisted = empNumber
    Else
        empNumber = empNumber + 1
        checkIndexExisted = checkIndexExisted(empInitial, empNumber)
    End If
    
End Function

Sub Main()
    MsgBox generateIndex("c")
End Sub
