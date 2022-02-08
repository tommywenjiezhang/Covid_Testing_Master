VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} testForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4452
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   6588
   OleObjectBlob   =   "testForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "testForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkIn_Click()
    Dim lastRow As Long
    Dim testType As String
    testType = ""
    Dim hasSymptom As String
    Dim db As New testDb
    Dim pcr As Boolean
    Dim rapid As Boolean
    Dim name As String
    Dim dobStr As String
    Dim dobDate As Date
    
    
    
    
    
    Dim empID As String
    empID = empList.Range("A" & ActiveCell.Row).value
    
    
    name = ActiveCell.value
    With testRoster
            lastRow = .Cells(.Rows.count, "A").End(xlUp).Row + 1
            .Range("A" & lastRow & ":Q" & .Rows.count).ClearContents
    End With
    
    If testForm.symptomChk.value = True Then
        hasSymptom = "Y"
    Else
        hasSymptom = "N"
    End If
    
     If Not testForm.dobTxt.value = "" Then
        On Error GoTo wrong_dob
        dobStr = validationHelper.birthdayExtract(testForm.dobTxt.value)
        dobDate = CDate(dobStr)
        db.updateBirthday dobDate, Trim(ActiveCell.value)
    End If

    
    If testForm.rapidTest.value = True And testForm.pcrChk = True Then
        db.insertTesting empID, name, Now, testForm.symptomChk.value, "RAPID"
        db.insertTesting empID, name, Now, testForm.symptomChk.value, "PCR"
        insertTest lastRow, "RAPID", empID, hasSymptom, dobDate
        insertTest (lastRow + 1), "PCR", empID, hasSymptom, dobDate
    ElseIf testForm.pcrChk.value = True Then
        db.insertTesting empID, name, Now, testForm.symptomChk.value, "PCR"
        insertTest lastRow, "PCR", empID, hasSymptom, dobDate
    ElseIf testForm.rapidTest.value = True Then
        db.insertTesting empID, name, Now, testForm.symptomChk.value, "RAPID"
        insertTest lastRow, "RAPID", empID, hasSymptom, dobDate
    End If
    
    testRoster.Cells.EntireColumn.AutoFit
    testRoster.Select
    testRoster.Cells(lastRow, "A").Select
   
done:
    Unload Me
    testRoster.Cells(lastRow, "A").Speak
    Exit Sub
    
wrong_dob:
    With testForm.dobTxt
        .SetFocus
        .BackColor = RGB(255, 255, 153)
    End With
    testForm.warning.Visible = True
    testForm.warning.Caption = "Please enter correct DOB"
    
End Sub

Private Sub insertTest(ByVal lastRow As Long, testType As String, empID As String, hasSymptom As String, dobDate As Date)
    
        With testRoster
                .Cells(lastRow, "A").value = empID
                .Cells(lastRow, "B").value = ActiveCell.value
                .Cells(lastRow, "C").value = Now
                .Cells(lastRow, "C").NumberFormat = "hh:mm:ss AM/PM"
                .Range("D" & lastRow).value = hasSymptom
                .Range("E" & lastRow).value = testType
                .Range("F" & lastRow).value = dobDate
        End With
End Sub

Private Sub closeBtn_Click()
Unload Me

End Sub

Private Sub populate_birthday()
    Dim last_row As Long
    Dim value As Variant
    Dim dob As String
    
    
    Dim empID As String
    empID = empList.Range("A" & ActiveCell.Row).value
    
    With empBirthday
        last_row = .Cells(.Rows.count, "A").End(xlUp).Row
        value = Application.VLookup(empID, .Range("A1:B" & last_row), 2, False)
        If Not IsError(value) Then
            dob = format(CDate(value), "mm/dd/yyyy")
            testForm.dobTxt.value = dob
            
        End If
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim empID As String
    empID = empList.Range("A" & ActiveCell.Row).value

    Me.empNameLal.Caption = "Testing for: " & ActiveCell.value & " : " & empID
    
    On Error Resume Next
    populate_birthday
        
End Sub
