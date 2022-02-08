Attribute VB_Name = "addNewNOTest"
Sub addNoTest()
     Dim util As New testUtil
    
    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, empList.Range("B2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
            noTestfrm.Show
        End If
    End If
End Sub
