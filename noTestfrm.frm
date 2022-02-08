VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} noTestfrm 
   Caption         =   "UserForm1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6720
   OleObjectBlob   =   "noTestfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "noTestfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addNoTestBtn_Click()
    Dim dateStr As String
    Dim empID As String
    Dim empName As String
    Dim currRow As Integer
    Dim db As New testDb
    Dim restartDate As Date
    
    If Not noTestfrm.noTestDate.value = "" Then
        dateStr = validationHelper.birthdayExtract(noTestfrm.noTestDate)
        currRow = ActiveCell.Row
        restartDate = CDate(dateStr)
        restartDate = DateAdd("m", 3, restartDate)

        With empList
            empID = .Cells(currRow, 1).value
            empName = .Cells(currRow, 2).value
            db.insertNoTest empID, empName, restartDate
        End With
        Unload Me
    Else
        MsgBox "Please enter the restart Test Date"
        noTestfrm.noTestDate.BackColor = RGB(255, 255, 0)
    End If
End Sub

Private Sub addOrDeleteCbo_Change()
    If Me.addOrDeleteCbo.value = "Add Employee to No Test List" Then
        Me.addFrame.Visible = True
        Me.deleteFrame.Visible = False
    Else
        Me.addFrame.Visible = False
        Me.deleteFrame.Visible = True
    End If

End Sub

Private Sub closeBtn_Click()
    Unload Me
End Sub

Private Sub deleteEmp_Click()
    Dim empID As String
    Dim empName As String
    Dim currRow As Integer
    Dim db As New testDb
    
    currRow = ActiveCell.Row
    
    With empList
        empID = .Cells(currRow, 1).value
        db.deleteNoTest empID
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.deleteFrame.Visible = False
    With Me.addOrDeleteCbo
        .AddItem "Add Employee to No Test List"
        .AddItem "Delete employee to No Test List"
    End With
End Sub
