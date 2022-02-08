VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} send_txt_frm 
   Caption         =   "UserForm2"
   ClientHeight    =   5220
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   6438
   OleObjectBlob   =   "send_txt_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "send_txt_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()
    Unload Me
End Sub

    
    

Private Sub btnSubmit_Click()
    If Me.phone_txt.value = "" Or Me.carrier_cbo.value = "" Then
            Me.phone_txt.BackColor = RGB(255, 255, 153)
            Me.carrier_cbo.BackColor = RGB(255, 255, 153)
            Me.warning_lbl.Visible = True
    Else
        Dim waitTime As Integer
        Dim msg As String
        Dim carrier As String
        Dim phone_num As String
        
        waitTime = CInt(Me.wait_time_txt.value)
        msg = Me.msg_txt.value
        carrier = Me.carrier_cbo.value
        phone_num = Me.phone_txt.value
        
        Call send_text.send_message(phone_num, carrier, msg, waitTime)
    End If
    
End Sub

Private Sub UserForm_Initialize()
    With carrier_cbo
    .List = Array("AT&T", "Sprint", "T-Mobile", "Verizon", "Boost", "Cricket", "Metro_PCS", "Tracfone", "US_Cellular", "Virgin_Mobile")
    End With
    
    With msg_txt
        .Text = "Your Test is resulted,Please come back to the testing room"
    End With
    
    With wait_time_txt
        .Text = "15"
    End With
    
    
End Sub
