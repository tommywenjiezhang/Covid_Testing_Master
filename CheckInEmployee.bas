Attribute VB_Name = "CheckInEmployee"
Option Explicit

Sub checkIn()
    
'Find the last used row in a Column: column A in this example
    Dim util As New testUtil
    
    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, empList.Range("B2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
            testForm.Show
        End If
    End If
End Sub

