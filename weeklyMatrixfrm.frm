VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} weeklyMatrixfrm 
   Caption         =   "Missing Test List"
   ClientHeight    =   6630
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   8460.001
   OleObjectBlob   =   "weeklyMatrixfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "weeklyMatrixfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
   With Me.testTypecbo
        .AddItem "RAPID"
        .AddItem "PCR"
    End With
    Me.endDateTxt.value = format(Date, "mm/dd/yyyy")
    Me.startDateTxt.value = format(DateAdd("d", -7, Date), "mm/dd/yyyy")
End Sub


Private Sub btnClose_Click()
 Unload Me
End Sub

Private Sub btnSumit_Click()
    Dim startDate  As String
    Dim endDate As String
    Dim typeOfTest As String
    
    
    If Not Me.endDateTxt.value = "" And Not Me.startDateTxt.value = "" Then
        startDate = validationHelper.birthdayExtract(Me.startDateTxt.value)
        endDate = validationHelper.birthdayExtract(Me.endDateTxt.value)
        getReport startDate, endDate
    Else
         Me.endDateTxt.BackColor = RGB(255, 255, 0)
            Me.startDateTxt.BackColor = RGB(255, 255, 0)
        
        
    End If
End Sub

Private Sub getReport(ByVal startDateStr As String, endDateStr As String)
    Dim startDate As Date
    Dim endDate As Date
    
     If Not IsError(CDate(startDateStr)) And Not IsError(CDate(endDateStr)) Then
        startDate = CDate(startDateStr)
        endDate = CDate(endDateStr)
        getData endDate, startDate, CStr(Me.testTypecbo.value)
        
    End If
    
End Sub



Private Sub getData(ByVal endDate As Date, ByVal startDate As Date, ByVal typeOfTest As String)
    Dim last_row As Long
    Dim new_sht As Worksheet
    Dim new_wb As Workbook
    Dim copy_rng As Range
    Dim idx As Long
    Dim db As New testDb
    Dim util As New testUtil
    Dim Result As Variant
    Dim data_sht As Worksheet
    Dim start As Long
    Dim vaccine_copy_rng As Range
    Dim dash_last_row As Long
    Dim filename As String
    Dim weekly_sht As Worksheet
    
    
    
    
    
    filename = "Missing Tests Weekly Report for " & format(startDate, "mm-dd-yy")
    
    Set new_wb = Workbooks.Add
    
    
    
    With empList
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        Set copy_rng = .Range("B2:B" & last_row)
        Set vaccine_copy_rng = .Range("E2:E" & last_row)
        
            Result = db.getWeeklyMissingTest(typeOfTest, startDate, endDate)
    
            Set data_sht = new_wb.Sheets(1)
            
            Dim j As Long
            If Not util.isArrayEmpty(Result) Then
                start = 2
                With data_sht
                    .Cells(1, 1).value = "empName"
                    .Cells(1, 2).value = "TestDate"
                    .Cells(1, 3).value = "Most recent Test"
                   
                End With
                For j = LBound(Result, 2) To UBound(Result, 2)
                    With data_sht
                        .Cells(start + j, 1).value = Result(0, j)
                        .Cells(start + j, 2).value = format(Result(1, j), "mm/dd/yyyy")
                        .Cells(start + j, 3).value = Result(2, j)
                    End With
                     
                Next j
            End If
            
            
        With new_wb
            Set new_sht = .Sheets.Add
            Set weekly_sht = .Sheets.Add(After:=.Sheets(.Sheets.count))
            getWeeklyFrequency weekly_sht, typeOfTest, startDate, endDate
            With new_sht
                .Range("A1").value = "empName"
                .Range("B1").value = "Vaccination Record"
                .Range("C1").value = "Most Recent Test"
                .Cells(1, 4).value = "Test"
                .Cells(1, 5).value = "requency"
                copy_rng.Copy
                .Range("A2").PasteSpecial xlPasteValues
                vaccine_copy_rng.Copy Destination:=.Range("B2")
                dash_last_row = .Cells(.Rows.count, 1).End(xlUp).Row
                
                Dim s As Long
                Dim lookup_value As Variant
                Dim weekly_lookup As Variant
                    For s = 2 To dash_last_row
                        lookup_value = Application.VLookup(.Cells(s, 1).value, data_sht.UsedRange, 2, False)
                        If Not IsError(lookup_value) Then
                                .Cells(s, 3).value = lookup_value
                                .Cells(s, 3).NumberFormat = "mm/dd/yyyy"
                                .Cells(s, 1).Interior.color = vbRed
                                .Cells(s, 4).value = typeOfTest
                           Else
                                .Cells(s, 3).value = ""
                                weekly_lookup = Application.VLookup(.Cells(s, 1).value, weekly_sht.UsedRange, 2, False)
                                If Not IsError(weekly_lookup) Then
                                    .Cells(s, 5).value = weekly_lookup
                                    If CInt(weekly_lookup) <= 1 Then
                                        If .Cells(s, 2).value = "No Vaccine" And .Cells(s, 4).value = "RAPID" Then
                                            .Cells(s, 5).Interior.color = vbMagenta
                                        End If
                                    End If
                                End If
                        End If
                    Next s
                
            End With
        End With
        
    End With
    
    generatePdf new_wb, startDate, filename & ".pdf", typeOfTest
    
End Sub

Private Sub fillMatrix(ByRef wb As Workbook)
    Dim last_row As Long, filterSht_lastRow As Long
    Dim idx As Long, dateCol As Long
    Dim dateStr As String
    Dim filteredRange As Variant
    Dim util As New testUtil
    Dim lookup_value As Variant
    
    

    
    With wb.Sheets(1)
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        
        For idx = 2 To last_row
            
            For dateCol = 2 To 8
                dateStr = .Cells(1, dateCol).value
                With wb.Sheets(2)
                    filterSht_lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
                    .Range("A1:C" & filterSht_lastRow).AutoFilter Field:=2, Criteria1:=DateSerial(Year(dateStr), Month(dateStr), Day(dateStr))
                    On Error Resume Next
                    filteredRange = .Range("A2:C" & filterSht_lastRow).SpecialCells(xlCellTypeVisible)
                    On Error GoTo 0
                    
                        With wb.Sheets(1)
                                lookup_value = Application.VLookup(.Cells(idx, 1).value, filteredRange, 2, False)
                                If Not IsError(lookup_value) Then
                                   .Cells(idx, dateCol).value = lookup_value
                                Else
                                   .Cells(idx, dateCol).value = ""
                             End If
                        End With
                    .ShowAllData
                End With
            Next dateCol
            
        Next idx
        
    End With
done:
 Exit Sub
 
no_cell_found:
    MsgBox "No record found"
End Sub


Private Sub generatePdf(ByRef wb As Workbook, ByVal startDate As Date, ByVal filepath As String, ByVal testType As String)
    Dim tfo As New TestExport
    Dim new_sht As Worksheet
    
    Set new_sht = wb.Sheets(1)

    new_sht.UsedRange.columns.AutoFit

    With new_sht.PageSetup
        .CenterHeader = "&B&20" & "Missing " & testType & " Report for " & format(startDate, "mm-dd-yy")
        .RightFooter = "Page: " & "&P"
        .CenterHorizontally = True
        .PrintArea = new_sht.UsedRange.Address
    End With

    On Error GoTo pdf_error
    new_sht.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        filename:=ThisWorkbook.path & "\" & filepath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        

done:
    Exit Sub
    
pdf_error:
    MsgBox "Cannot create PDF report"
    
End Sub


Sub getWeeklyFrequency(ByRef sht As Worksheet, typeOfTest As String, startDate As Date, endDate As Date)
    Dim Result As Variant
    Dim db As New testDb
    Dim util As New testUtil
    
    Result = db.getAggragateWeeklyTest(typeOfTest, startDate, endDate)
    
    Dim j As Long
            If Not util.isArrayEmpty(Result) Then
                start = 2
                With sht
                    .Cells(1, 1).value = "empName"
                    .Cells(1, 2).value = "Number of Test"
                End With
                For j = LBound(Result, 2) To UBound(Result, 2)
                    With sht
                        .Cells(start + j, 1).value = Result(0, j)
                        .Cells(start + j, 2).value = Result(1, j)
                    End With
                     
                Next j
            End If
    
End Sub
