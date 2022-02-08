Attribute VB_Name = "getTesting"
Sub getTestingByEmp()

Dim util As New testUtil

If util.InRange(ActiveCell, empList.Range("B2:B1000")) Then
    testLookUpfrm.Show
    
Else

    MsgBox "Wrong Area, Please select a name from employee list"

End If



End Sub
