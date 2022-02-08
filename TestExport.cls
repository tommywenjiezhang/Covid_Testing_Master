VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private employeeTest As Dictionary
Private testWbDir As String
Private empHistory As String
Private todayFolder As String
Private exportFileName  As String

Private Sub Class_Initialize()
    Set employeeTest = New Dictionary
    testWbDir = "\test_history"
    empHistory = "\employee_history"
    todayFolder = format(Date, "mmmm_dd_yyyy")
End Sub

Property Get full_path() As String
    full_path = ThisWorkbook.path & testWbDir & "\" & todayFolder
End Property

Property Get employee_path() As String
    employee_path = ThisWorkbook.path & empHistory & "\" & todayFolder
End Property

Public Sub Add(name As String, testDate As Date)
    employeeTest.Add name, Date
End Sub

Public Sub makefolder()
    If Not FolderExists(Me.full_path) Then
        FolderCreate Me.full_path
    End If
    If Not FolderExists(Me.full_path & "\pdf") Then
        FolderCreate Me.full_path & "\pdf"
    End If
    If Not FolderExists(Me.employee_path) Then
        FolderCreate Me.employee_path
    End If
End Sub

Function FolderCreate(ByVal spath As String) As Boolean

  'create full sPath at once, if required
  'returns False if folder does not exist and could NOT be created, True otherwise
  'sample usage: If CreateFolder("C:\toto\test\test") Then debug.print "OK"

    Dim fs As Object
    Dim FolderArray
    Dim Folder As String, i As Integer, sShare As String

    If Right(spath, 1) = "\" Then spath = Left(spath, Len(spath) - 1)
    Set fs = CreateObject("Scripting.FileSystemObject")
    'UNC path ? change 3 "\" into 3 "@"
    If spath Like "\\*\*" Then
        spath = Replace(spath, "\", "@", 1, 3)
    End If
    'now split
    FolderArray = Split(spath, "\")
    'then set back the @ into \ in item 0 of array
    FolderArray(0) = Replace(FolderArray(0), "@", "\", 1, 3)
    On Error GoTo hell
    'start from root to end, creating what needs to be
    For i = 0 To UBound(FolderArray) Step 1
        Folder = Folder & FolderArray(i) & "\"
        If Not fs.FolderExists(Folder) Then
            fs.CreateFolder (Folder)
        End If
    Next
    FolderCreate = True
hell:
End Function


Function FolderExists(ByVal path As String) As Boolean
Dim isExist As Boolean
FolderExists = False
Dim fso As New FileSystemObject

isExist = fso.FolderExists(path)

If fso.FolderExists(path) Then FolderExists = True

End Function

Public Sub exportFile()
    Dim new_wb As Workbook
    Dim last_row As Long
    makefolder
    If Dir(Me.full_path & exportFileName) = "" Then
    
        Set new_wb = Workbooks.Add
    
        With new_wb.Sheets(1)
            .Range("A1").value = "Employee Name"
            .Range("B1").value = "Date"
            .Range("C1").value = "Export Time"
            .Range("A2:A" & UBound(employeeTest.Keys) + 1) = WorksheetFunction.Transpose(employeeTest.Keys)
            .Range("B2:B" & UBound(employeeTest.Items) + 1) = WorksheetFunction.Transpose(employeeTest.Items)
            .Range("C2:C" & UBound(employeeTest.Items) + 1).value = Now
            .columns("A:C").AutoFit
            .SaveAs filename:=Me.full_path & exportFileName
        End With
    
    
        new_wb.Close SaveChanges:=True
    Else
        Set new_wb = Workbooks.Open(Me.full_path & exportFileName)
        With new_wb.Sheets(1)
            last_row = .Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Row
            .Range("A" & last_row & ":A" & last_row + UBound(employeeTest.Keys) + 1) = WorksheetFunction.Transpose(employeeTest.Keys)
            .Range("B" & last_row & ":B" & last_row + UBound(employeeTest.Items) + 1) = WorksheetFunction.Transpose(employeeTest.Items)
            .Range("C" & last_row & ":C" & last_row + UBound(employeeTest.Items) + 1).value = Now
            .columns("A:C").AutoFit
    
        End With
        new_wb.Close SaveChanges:=True
    End If
    
End Sub






