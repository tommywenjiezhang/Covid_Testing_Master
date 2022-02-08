Attribute VB_Name = "Module1"
Sub prefillTesting()
    Dim last_row As Long
    Dim idx As Long
    With testRoster
        last_row = .Cells(Rows.count, "A").End(xlUp).Row
        For idx = 3 To last_row
            If IsEmpty(.Cells(idx, "G")) Then
                .Cells(idx, "G").value = "N"
            End If
        Next idx
        .Cells(last_row, "A").EntireRow.Select
    End With
End Sub
