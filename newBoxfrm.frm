VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newBoxfrm 
   Caption         =   "Add New Box"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "newBoxfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newBoxfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub submit_Click()
    Dim box_num As Long
    Dim expireDate As Date
    Dim dateHolder As Variant
    Dim lot_Number As String
    
    If Not newBoxfrm.boxNumber.value = "" And _
        Not newBoxfrm.expireDate.value = "" And _
        newBoxfrm.lotNumber.value = "" Then
    
        box_num = CLng(newBoxfrm.boxNumber.value)
        dateHolder = validationHelper.birthdayExtract(newBoxfrm.expireDate.value)
        expireDate = CDate(dateHolder)
        lot_Number = CStr(newBoxfrm.lotNumber.value)
        
    Else
        With newBoxfrm
            .boxNumber.BackColor = RGB(255, 255, 0)
            .expireDate.BackColor = RGB(255, 255, 0)
            .lotNumber.BackColor = RGB(255, 255, 0)
            .warning.Visible = True
        End With
        
    
    
    End If
    
    
End Sub

