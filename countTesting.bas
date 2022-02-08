Attribute VB_Name = "countTesting"
Sub countTotal()
Dim employeeTotal As Long
Dim visitorTotal As Long
Dim total As Long
Dim rapidTestTotal As Long
Dim pcrTestTotal As Long




With testRoster
    employeeTotal = .Cells(.Rows.count, 1).End(xlUp).Row - 2

End With

rapidTestTotal = 0
Dim idx As Long

For idx = 3 To (employeeTotal + 2)
    With testRoster
        If Not IsEmpty(.Cells(idx, "E").value) And .Cells(idx, "E").value = "RAPID" Then
            rapidTestTotal = rapidTestTotal + 1
        End If
    End With
Next idx


pcrTestTotal = employeeTotal - rapidTestTotal



With visitorTesting
    visitorTotal = .Cells(.Rows.count, 1).End(xlUp).Row - 2

End With


total = employeeTotal + visitorTotal

MsgBox "Total TESTED: " & total & vbCrLf & "Empolyee Testing: " & employeeTotal & " (PCR: " & _
CStr(pcrTestTotal) & _
", RAPID: " & CStr(rapidTestTotal) & ") " & _
vbCrLf & "Visitor testing: " & visitorTotal

End Sub
