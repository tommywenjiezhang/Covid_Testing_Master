Attribute VB_Name = "send_text"
Sub send_message(phone_num As String, phone_carrier As String, msg As String, waitTime As Integer)
    
    Dim name As String
    Dim execute_str As String
    
    
    
    Dim path As String
    name = ActiveCell.value

    If name = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        
        execute_str = "D:\programs\send_message\dist\send_text_message.exe " & _
        " --name " & Chr(34) & name & Chr(34) & _
        " --phone " & Chr(34) & phone_num & Chr(34) & _
        " --msg " & Chr(34) & msg & Chr(34) & _
        " --carrier " & Chr(34) & phone_carrier & Chr(34) & _
        " --waitTime " & waitTime
        
        Debug.Print execute_str
        obj = Shell(execute_str, vbMinimizedFocus)
    End If
    MsgBox "Text is being sent"
End Sub

