VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} visitorfrm 
   Caption         =   "UserForm1"
   ClientHeight    =   4230
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   7878
   OleObjectBlob   =   "visitorfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "visitorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkIn_Click()
    Dim visitorName As String
    Dim visitorDOB As String
    Dim testType As String
    Dim hasSymptom As String
    Dim lastRow As Long
    Dim dobDate As Date
    Dim db As New visitorTestingDb
    
    
     
    If visitorfrm.rapidTest.value = True Then
        testType = testType & "RAPID"
    End If
    
    If visitorfrm.rapidTest.value = True And visitorfrm.pcrChk = True Then
        testType = testType & "&"
       
    End If
    
    If visitorfrm.pcrChk.value = True Then
        testType = testType & "PCR"
    End If
    
    If visitorfrm.symptomChk.value = True Then
        hasSymptom = "Y"
    Else
        hasSymptom = "N"
    End If
    
    If visitorfrm.visitorName.value = "" Or visitorfrm.birthday.value = "" Then
        With visitorfrm
            .visitorName.BackColor = RGB(255, 255, 102)
            .warning.Visible = True
        End With
    
    Else
        visitorTesting.Activate
        With visitorTesting
            lastRow = .Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Row
            visitorTesting.Cells(lastRow, 1).value = UCase(visitorfrm.visitorName.value)
            dobDate = CDate(validationHelper.birthdayExtract(visitorfrm.birthday.value))
            .Cells(lastRow, 2).value = Now
            .Cells(lastRow, 2).NumberFormat = "hh:mm AM/PM"
            .Cells(lastRow, 3).value = hasSymptom
            .Cells(lastRow, 4).value = testType
            .Cells(lastRow, 5).value = dobDate
            .Cells(lastRow, 5).NumberFormat = "mm/dd/yyyy"
            .Cells.EntireColumn.AutoFit
            .columns("F").ColumnWidth = 45
            .Range(.Cells(lastRow, 1), .Cells(lastRow, "F")).Select
        End With
        
        db.insertTesting UCase(visitorfrm.visitorName.value), Now, visitorfrm.symptomChk.value, testType, dobDate
        Unload Me
    End If
    
    
End Sub

Private Sub closeBtn_Click()
Unload Me

End Sub

