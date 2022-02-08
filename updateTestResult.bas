Attribute VB_Name = "updateTestResult"
Sub updateTestResult()
    Dim db As New testDb
    
    Dim lastRow As Long
    Dim idx As Long
    Dim empID As String
    Dim Result As String
    
    Dim message As String
    With testRoster
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For idx = 3 To lastRow
            If IsEmpty(.Cells(idx, "G")) Then
                .Cells(idx, "G").Interior.color = RGB(255, 255, 102)
                message = "Some result not filled, please fill out the result and export again"
            Else
                If Not IsEmpty(.Cells(idx, "E")) Then
                    empID = Trim(.Cells(idx, 1).value)
                    If checkIfbothTest(.Cells(idx, "E").value) Then
                        Result = UCase(Left(.Cells(idx, "G").value, 1))
                        db.updateTestResult empID, Now, "RAPID", Result
                        db.updateTestResult empID, Now, "PCR", Result
                    Else
                        db.updateTestResult empID, Now, .Cells(idx, "E").value, Result
                    End If
                    
                End If
            End If
        Next idx
    End With
    
    If Not message = "" Then
        MsgBox message
        testRoster.Activate
    End If
    
End Sub

Function checkIfbothTest(ByVal str As String) As Boolean
    If InStr(str, "&") > 0 Then
        checkIfbothTest = True
    Else
        checkIfbothTest = False
        
    End If
    
End Function
