Attribute VB_Name = "clearTesting"
Option Explicit

Sub clearTesting()

Dim answer  As Integer


answer = MsgBox("Are you sure to clear all the testing information ", vbQuestion + vbYesNo)
    If answer = vbYes Then
    
        With testRoster
            .Range(.Cells(3, 1), .Cells(.Rows.count, "G")).ClearContents
            .Range(.Cells(3, 1), .Cells(.Rows.count, "G")).Interior.ColorIndex = 0
            
        End With
        
        With visitorTesting
            .Range(.Cells(3, 1), .Cells(.Rows.count, "F")).ClearContents
        End With
    Else
    Exit Sub

    End If


End Sub
