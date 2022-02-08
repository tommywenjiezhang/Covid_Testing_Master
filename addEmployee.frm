VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addEmployee 
   Caption         =   "Add Employee"
   ClientHeight    =   4200
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   5538
   OleObjectBlob   =   "addEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim empID As String




Private Sub addNameBtn_Click()
    Dim firstName As String
    Dim lastName As String
    Dim db As New testDb
    Dim index As Long
    Dim fullName As String
    Dim first_initial As String
    Dim dobtext As String, dobDate As Date
    Dim lastRow  As Long


    
    If Me.firstNameTxt.value = "" Or Me.lastNameTxt.value = "" Then
        Me.firstNameTxt.BackColor = RGB(255, 255, 153)
        Me.lastNameTxt.BackColor = RGB(255, 255, 153)
        Me.warning.Visible = True
    
    Else
    
        firstName = Me.firstNameTxt.value
        lastName = Me.lastNameTxt.value
        fullName = UCase(lastName) & "," & UCase(firstName)
        first_initial = UCase(Left(Me.lastNameTxt.value, 1))
        index = generateIndex.generateIndex(first_initial) + 1
        empID = first_initial & CStr(index)
        
        If Not Me.dobTxt.value = "" Then
            dobtext = validationHelper.birthdayExtract(Me.dobTxt.value)
            dobDate = CDate(dobtext)
            db.insertEmpName fullName, empID, dobDate
        Else
        
            db.insertEmpName fullName, empID
        
        End If
        last_row = empList.Cells(empList.Rows.count, 1).End(xlUp).Offset(1, 0).Row
        empList.Cells(last_row, 2).value = UCase(lastName) & "," & UCase(firstName)
        empList.Cells(last_row, 1).value = empID
        empList.Cells(last_row, 1).Select
        
        Unload Me
    End If
    
    
    
    
    
    
    
    
   
    
  
    
    
End Sub

Private Sub closeAddNameBtn_Click()
    Unload Me
End Sub

Private Sub lastNameTxt_Change()
    Dim index As Long
    first_initial = UCase(Left(Me.lastNameTxt.value, 1))
    index = generateIndex.generateIndex(first_initial)
    empID = first_initial & CStr(index)
    With Me.newEmpID
    .Caption = "New Employee ID: " & first_initial & CStr(index)
    .font.Bold = True
    .font.Size = 14
    End With
    
End Sub

Private Sub UserForm_Initialize()
    Dim id As String
    
End Sub
