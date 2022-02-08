Attribute VB_Name = "SendReport"
Sub send_instant_email()
    Dim email As String
    Dim execute_str As String
    
    Dim util As New testUtil
    
    
    Dim path As String

    execute_str = "D:\programs\python_32\python.exe -i " & "D:\programs\sendEmail\runner.py --d False"
    obj = Shell(execute_str, vbMinimizedFocus)
    
    Debug.Print execute_str
    MsgBox "sending email"
End Sub
