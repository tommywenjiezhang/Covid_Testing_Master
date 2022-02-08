Attribute VB_Name = "undoLast"
Sub undoLast()

    Dim last_row As Long
    If ActiveSheet.name = "Visitor" Or ActiveSheet.name = "Test Roster" Then
            With ActiveSheet
                last_row = .Cells(.Rows.count, 1).End(xlUp).Row
                If last_row > 2 Then
                    .Range(.Cells(last_row, "A"), .Cells(last_row, "G")).ClearContents
                End If
            End With
    End If
    
    

End Sub
