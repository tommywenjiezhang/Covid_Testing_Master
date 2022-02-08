Attribute VB_Name = "importTest"
Option Explicit



Sub importTest()
    Dim Result As Variant
    Dim db As New testDb
    Dim todayDate As Date
    Dim tenDayAgo As Date
    Dim util As New testUtil
    
    todayDate = Date
    tenDayAgo = DateAdd("d", -7, Date)
    
    testImport.Cells.ClearContents
    
    Dim i As Long
    Dim j As Long
    
    Dim start As Long
    start = 1
    Result = db.getTestHistory(tenDayAgo, todayDate)
    If Not util.isArrayEmpty(Result) Then
        For j = LBound(Result, 2) To UBound(Result, 2)
            With testImport
                .Cells(start + j, 1).value = Result(0, j)
                .Cells(start + j, 2).value = Result(1, j)
                .Cells(start + j, 3).value = Result(2, j)
            End With
        Next j
    End If
    
    
    
    
End Sub

Sub importNoTestList()
    Dim Result As Variant
    Dim db As New testDb
    Dim util As New testUtil
    
    Dim i As Long
    Dim j As Long
    
    noTest.Cells.ClearContents
    Dim start As Long
    start = 1
    Result = db.getNoTestList()
    If Not util.isArrayEmpty(Result) Then
        For j = LBound(Result, 2) To UBound(Result, 2)
            With noTest
                .Cells(start + j, 1).value = Result(1, j)
                .Cells(start + j, 2).value = Result(2, j)
                .Cells(start + j, 3).value = format(Result(3, j), "mm/dd/yyyy")
            End With
        Next j
    End If
    
End Sub

Sub populate_testing()
    Dim last_row As Long
    Dim idx As Long
    Dim lookup_result As Variant
    Dim lookup_value As String
    Dim days As Long
    Dim message
    Dim testFrequency As Long
    Dim rapid_look_up_range As Variant, pcr_look_up_range As Variant, pcr_lookup_result As Variant, rapid_lookup_result As Variant
    Dim result_test As Variant
    
    testFrequency = CLng(empList.Range("F2").value)
    
    empList.Unprotect
    
     pcr_look_up_range = getFilterRange("PCR")
     rapid_look_up_range = getFilterRange("RAPID")
    
    
    With empList
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range(.Cells(2, 1), .Cells(last_row, 4)).Interior.ColorIndex = xlNone
        
    End With
    For idx = 2 To last_row
        If Not empList.Cells(idx, 1).value = "" Then
            lookup_value = empList.Cells(idx, 1).value
            pcr_lookup_result = Application.VLookup(lookup_value, pcr_look_up_range, 2, False)
            rapid_lookup_result = Application.VLookup(lookup_value, rapid_look_up_range, 2, False)
            If Not IsError(pcr_lookup_result) Then
                empList.Cells(idx, 3).value = CDate(pcr_lookup_result)
                empList.Cells(idx, 3).NumberFormat = "dddd, mm/dd/yy"
            Else
                 empList.Cells(idx, 3).value = "Test Not Found"
                 empList.Cells(idx, 3).Interior.color = RGB(255, 69, 0)
            End If
            If Not IsError(rapid_lookup_result) Then
                empList.Cells(idx, 4).value = CDate(rapid_lookup_result)
                empList.Cells(idx, 4).NumberFormat = "dddd, mm/dd/yy"
            Else
                empList.Cells(idx, 4).Interior.color = RGB(255, 69, 0)
                 empList.Cells(idx, 4).value = "Test Not Found"
            End If
        End If
    Next idx
    
    empList.Protect
    
End Sub


Function getFilterRange(ByVal filterType As String) As Variant
    Dim last_row As Long
    Dim filter_rng As Range
    
    Dim row_rng As Range
    
    
    With testImport
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range("A1").AutoFilter Field:=3, Criteria1:=filterType
        getFilterRange = Arr_Visible_Cells()
        
    End With
    
   
     
End Function


Function Arr_Visible_Cells() As Variant
Dim rRow As Range
Dim aArr() As Variant
Dim i As Long
Dim lCount As Long
Dim CellCount As Variant
Dim Range_To_Get As Variant

CellCount = testImport.UsedRange.Rows.count
Range_To_Get = testImport.UsedRange.Address

ReDim aArr(1 To 3, 1 To CellCount)

lCount = 1
i = 1
For Each rRow In testImport.Range(Range_To_Get)
    If lCount = 4 Then
    i = i + 1
    lCount = 1
    End If
        If rRow.Rows.Hidden = False Then
         aArr(lCount, i) = rRow
        Else
         GoTo Devo:
        End If
lCount = lCount + 1
Devo:
Next

ReDim Preserve aArr(1 To 3, 1 To i)

Arr_Visible_Cells = Application.Transpose(aArr)
End Function

Sub refreshNoTest()
    Dim idx As Long
    Dim look_rng As Range
    Dim last_row As Integer
    
    Set look_rng = noTest.UsedRange
    
    With empList
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        .Unprotect
    End With
    
    Dim lookup_value As Variant
    Dim lookup_result As Variant
    
    For idx = 2 To last_row
        If Not empList.Cells(idx, 1).value = "" Then
            lookup_value = empList.Cells(idx, 1).value
            lookup_result = Application.VLookup(lookup_value, look_rng, 3, False)
            
            If Not IsError(lookup_result) Then
                If CDate(lookup_result) >= Date Then
                    With Sheets("EMPLOYEE")
                        .Range("A" & idx).Interior.color = RGB(255, 0, 0)
                    End With
                End If
            End If
        End If
    Next idx
    
    empList.Protect

End Sub

Sub refresh_today_data()
    Dim emp_result As Variant
    Dim visitor_result As Variant
    Dim db As Variant
    Dim startDate As Date
    Dim endDate As Date
    Dim util As New testUtil
    Dim new_wb As Workbook
    Dim filename As String
    
    
    startDate = Date
    endDate = DateAdd("d", 1, startDate)
    
    
    
    Dim i As Long
    Dim j As Long
    
    Dim start As Long
    start = 3
        
    If ActiveSheet.CodeName = "testRoster" Then
        Set db = New testDb
        emp_result = db.getTestHistory(startDate, endDate)
    Else
        Set db = New visitorTestingDb
        visitor_result = db.getTestHistory(startDate, endDate)
    End If
    
    
        If Not util.isArrayEmpty(emp_result) Then
            Call clearTesting.clearTesting
            For j = LBound(emp_result, 2) To UBound(emp_result, 2)
                With testRoster
                    .Cells(start + j, 1).value = emp_result(0, j)
                    .Cells(start + j, 2).value = emp_result(3, j)
                    .Cells(start + j, 3).value = emp_result(1, j)
                    .Cells(start + j, 4).value = emp_result(2, j)
                End With
                 
            Next j
    End If

done:
    Exit Sub
End Sub


Sub Import_test_main()
    importTest
    populate_testing
    importNoTestList
    refreshNoTest
    Call refreshRoster.importBirthday
    Call refreshRoster.importVaccine
    Call refreshRoster.lookupVaccine
End Sub
