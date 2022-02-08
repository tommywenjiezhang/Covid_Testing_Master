VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePickerForm 
   Caption         =   "Date picker"
   ClientHeight    =   4755
   ClientLeft      =   48
   ClientTop       =   378
   ClientWidth     =   3762
   OleObjectBlob   =   "DatePickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private WithEvents Calendar1 As cCalendar
Attribute Calendar1.VB_VarHelpID = -1

Public Target As Range

Private Sub nextMonth_Click()
 Calendar1.nextMonth
End Sub

Private Sub previousMonth_Click()
    Calendar1.previousMonth
End Sub

Private Sub UserForm_Initialize()
    If Calendar1 Is Nothing Then
        Set Calendar1 = New cCalendar
        With Calendar1
            .Add_Calendar_into_Frame Me.Frame1
            .UseDefaultBackColors = False
            .DayLength = 3
            .MonthLength = mlENShort
            .Height = 140
            .Width = 180
            .GridFont.Size = 7
            .DayFont.Size = 7
            .Refresh
        End With
        Me.Height = 210 'Win7 Aero
        Me.Width = 193
        Me.Caption = "Screening History for "
        With Me.annoationLbl
            .Left = 60
            .Top = 155
        End With
        With Me.previousMonth
            .Left = 10
            .Top = 155
            .Caption = "<"
        
        End With
        
        With Me.nextMonth
            .Left = 150
            .Top = 155
            .Caption = ">"
        End With
        
    End If
End Sub

Public Property Get Calendar() As cCalendar
    Set Calendar = Calendar1
End Property



Public Sub MoveToTarget()
    Dim dLeft As Double, dTop As Double

    dLeft = Target.Left - ActiveWindow.VisibleRange.Left + ActiveWindow.Left
    If dLeft > Application.Width - Me.Width Then
        dLeft = Application.Width - Me.Width
    End If
    dLeft = dLeft + Application.Left
    
    dTop = Target.Top - ActiveWindow.VisibleRange.Top + ActiveWindow.Top
    If dTop > Application.Height - Me.Height Then
        dTop = Application.Height - Me.Height
    End If
    dTop = dTop + Application.Top
    
    Me.Left = IIf(dLeft > 0, dLeft, 0)
    Me.Top = IIf(dTop > 0, dTop, 0)
End Sub

Private Sub Calendar1_Click()
    Call CloseDatePicker(True)
End Sub

Private Sub Calendar1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call CloseDatePicker(False)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        CloseDatePicker (False)
    End If
End Sub

Sub CloseDatePicker(Save As Boolean)
    If Save And Not Target Is Nothing And IsDate(Calendar1.value) Then
        Target.value = Calendar1.value
    End If
    Set Target = Nothing
    Me.Hide
End Sub
