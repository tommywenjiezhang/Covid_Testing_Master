Attribute VB_Name = "Module4"
Sub test_tb_select_qry()
    Dim qb As New Query_Select_Builder
    
    
    qb.fromTable = "Testing"
    
    qb.SelectCol "empName", "empID", "timeIn", "typeOfTest"
    qb.Where = "typeOfTest = 'RAPID'"
    
    Debug.Print qb.GetQuery()
    
End Sub

Sub test_insert_qry()
    Dim db As New testDb
    Dim today As Date
    Dim arr As Variant
    
    today = Date
    arr = db.getTestHistory(DateAdd("m", -1, today), today)
End Sub
