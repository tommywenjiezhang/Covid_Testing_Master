Attribute VB_Name = "exportTesting"
Option Explicit

Sub exportTesting()
    Dim tfo As New TestExport
    Call updateTestResult.updateTestResult
    tfo.makefolder
    exportForm.Show
    
    
    
End Sub
