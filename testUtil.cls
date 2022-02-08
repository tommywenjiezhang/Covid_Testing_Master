VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function InRange(Range1 As Range, Range2 As Range) As Boolean
' returns True if Range1 is within Range2
Dim InterSectRange As Range
Set InterSectRange = Application.Intersect(Range1, Range2)
InRange = Not InterSectRange Is Nothing
Set InterSectRange = Nothing
End Function


Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Function isArrayEmpty(arr As Variant) As Boolean
    If Not IsArray(arr) Then
        isArrayEmpty = True
    ElseIf IsError(LBound(arr, 1)) And _
                             LBound(arr, 1) <= UBound(arr, 1) Then
        isArrayEmpty = True
    Else
        isArrayEmpty = False
    End If
End Function
