VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} testLookUpfrm 
   Caption         =   "Test Lookup"
   ClientHeight    =   3072
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   4620
   OleObjectBlob   =   "testLookUpfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "testLookUpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    Dim typeOfTest As String
    Dim db As New testDb
    Dim dateArr As Variant
    Dim util As New testUtil
    Dim empName As String
    
    
    typeOfTest = testLookUpfrm.testTypecbo.value
    empName = Trim(ActiveCell.value)
        
    
    dateArr = db.getTestedByEmp(empName, typeOfTest)

    
    If util.isArrayEmpty(dateArr) Then
        MsgBox "No Screening in the last five days"
    
    Else
    
        Dim j As Long
        Dim fDates As Variant
        ReDim fDates(LBound(dateArr, 2) To UBound(dateArr, 2) + 1)
        For j = LBound(dateArr, 2) To UBound(dateArr, 2)
            fDates(j) = format(CDate(dateArr(1, j)), "mm/dd/yyyy")
        Next j
        
        
        Call DatePickerForm.Calendar.ClearAllColoredDateArrays
        Call DatePickerForm.Calendar.AddColoredDateArray(RGB(34, 139, 34), fDates, True)
        With DatePickerForm
            .Caption = "Testing History for " & empName
            .annoationLbl.Caption = "Tested"
            .annoationLbl.BackColor = RGB(200, 255, 200)
            .annoationLbl.TextAlign = fmTextAlignCenter
        End With
        
        DatePickerForm.Show vbModal
        Cancel = True

        Erase fDates
        
    End If
End Sub

Private Sub UserForm_Initialize()
    With Me.testTypecbo
        .AddItem "RAPID"
        .AddItem "PCR"
    End With
End Sub
