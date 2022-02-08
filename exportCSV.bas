Attribute VB_Name = "exportCSV"
Function save_as_csv(strShift As String, Optional save_pdf As Boolean = False, Optional ByVal input_date As Date)
    Dim tfo As New TestExport
    Dim empFilename As String, vistorFileName As String
    Dim employee As Workbook, vistor As Workbook
    Dim emp_count As Long, vist_count As Long
    Dim empSht As Worksheet, vist_sht As Worksheet
    
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Not IsNull(input_date) Then
        todayDate = format(input_date, "YYYY-MM-DD")
    Else
        todayDate = format(Now, "YYYY-MM-DD")
    End If
    
    
    empFilename = tfo.full_path & "\" & todayDate & " " & strShift & "-emp-testing.xlsx"
    vistorFileName = tfo.full_path & "\" & todayDate & " " & strShift & "-vistor-testing.xlsx"
    
    Set employee = Workbooks.Add
    Set vistor = Workbooks.Add
    
    
    
    
    Set empSht = employee.Sheets(1)
    Set vist_sht = vistor.Sheets(1)
    
    
    getReport empSht, "EMPLOYEE", input_date
    getReport vist_sht, "VISITOR", input_date
    
    With empSht
        emp_count = .Cells(.Rows.count, 1).End(xlUp).Row - 2
        If emp_count < 0 Then
            emp_count = 0
        End If
    End With
    
    With vist_sht
        vist_count = .Cells(.Rows.count, 1).End(xlUp).Row - 2
        If vist_count < 0 Then
            vist_count = 0
        End If
    End With
    
    createPivot employee, "EMPLOYEE"
    createPivot vistor, "VISITOR"
    
    
    With empSht
        .Range("A1").value = "Empolyee Testing " & todayDate
        .Cells(.Rows.count, "A").End(xlUp).Offset(1, 0).value = "Total"
        .Cells(.Rows.count, "B").End(xlUp).Offset(1, 0).value = emp_count
        .SaveAs filename:=empFilename
    End With
    
    
    With vist_sht
        .Range("A1").value = "Visitor Testing " & todayDate
        .Cells(.Rows.count, "A").End(xlUp).Offset(1, 0).value = "Total"
        .Cells(.Rows.count, "B").End(xlUp).Offset(1, 0).value = vist_count
        .SaveAs filename:=vistorFileName
    End With
    
    
    reportFormat empSht
    reportFormat vist_sht
    
    If save_pdf Then
        Shell "taskkill /IM ""AcroRd32.exe"" /F"
        save_as_pdf strShift, empSht, "EMPLOYEE", input_date
        save_as_pdf strShift, vist_sht, "VISITOR", input_date
    End If
    
    employee.Close
    vistor.Close
    
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Your CSV file was exported to " & empFilename, vbInformation
    
done:
    MsgBox "Total Testing for " & strShift & " shift : " & CStr((emp_count + vist_count))
    
    
    Exit Function
    
End Function



Sub getReport(ByRef data_sht As Worksheet, ByVal reportType As String, Optional ByVal inputDate As Date)
    Dim Result As Variant
    Dim db As Variant
    Dim startDate As Date
    Dim endDate As Date
    Dim util As New testUtil
    Dim new_wb As Workbook
    Dim filename As String
    
    
    
    If IsNull(inputDate) Then
    
        startDate = Date
        endDate = DateAdd("d", 1, startDate)
       
    Else
        
        startDate = inputDate
        endDate = DateAdd("d", 1, inputDate)
    End If
    
    
    
    Dim i As Long
    Dim j As Long
    
    Dim start As Long
    start = 3
        
    If reportType = "EMPLOYEE" Then
        Set db = New testDb
        Result = db.getTestHistory(startDate, endDate, True)
    Else
        Set db = New visitorTestingDb
        Result = db.getTestHistory(startDate, endDate)
    End If
    
    
        If Not util.isArrayEmpty(Result) Then
            With data_sht
                .Cells(2, 1).value = "Name"
                .Cells(2, 2).value = "TestDate"
                .Cells(2, 3).value = "typeOfTest"
            End With
            
            For j = LBound(Result, 2) To UBound(Result, 2)
                With data_sht
                    .Cells(start + j, 1).value = Result(3, j)
                    .Cells(start + j, 2).value = Result(1, j)
                    .Cells(start + j, 3).value = Result(2, j)
                End With
                 
            Next j
    End If

done:
    Exit Sub
End Sub



Sub reportFormat(ByRef sht As Worksheet)
    Dim last_row As Long
    Dim table_rng As Range
    With sht
        .Range("A1:C1").Merge
        With .Range("A1:C1").font
            .name = "Calibri"
            .Size = 16
            .Bold = True
        End With
        .columns("A:C").EntireColumn.AutoFit
        last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        Set table_rng = .Range("A2:C" & last_row)
        .ListObjects.Add(xlSrcRange, table_rng, , xlYes).name = _
        "Table1"
        .ListObjects("Table1").TableStyle = "TableStyleMedium2"
    End With
    
End Sub


Private Sub createPivot(ByRef wb As Workbook, reportType As String)
    Dim data_sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    Dim pvt_sht As Worksheet
    Dim data_rng As Range
    Dim pi As PivotItem
    Dim pcr_count As Long
    Dim rapid_count As Long
    Dim data_last_row As Long
    
    
    With wb
        Set data_sht = .Sheets(1)
        With data_sht
            data_last_row = .Cells(.Rows.count, 1).End(xlUp).Row
        End With
        
        Set data_rng = .Sheets(1).UsedRange
        Set pvtCache = .PivotCaches.Create _
        (SourceType:=xlDatabase, SourceData:=data_rng)
    End With
    
    
    If data_last_row > 2 Then
        Set pvt = pvtCache.createPivotTable(TableDestination:=data_sht.Cells(2, "D"), TableName:="TestingReport")
        
          With pvt.PivotFields("typeOfTest")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        With pvt.PivotFields("typeOfTest")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlCount
        End With
    
    End If
End Sub



Sub save_as_pdf(spath As String, sht As Worksheet, reportType As String, Optional ByVal input_date As Date)

Dim tfo As New TestExport

Dim todayDate As String
Dim filepath As String
Dim last_row As Long
Dim file_save_location As String



With sht
    last_row = .Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Row
    .Cells.EntireColumn.AutoFit
    .Range("G3:G" & last_row).Interior.ColorIndex = 0
    .Range("D2").value = "Type of Test"
End With

    If Not IsNull(input_date) Then
        todayDate = format(input_date, "dddd dd mmm, yyyy")
        filepath = format(input_date, "mm-dd-yy")

    Else
        todayDate = format(Now, "dddd dd mmm, yyyy")
        filepath = format(Now, "mm-dd-yy")
    End If



file_save_location = tfo.full_path & "\pdf\" & filepath & spath & " " & reportType & "_testing.pdf"


With sht.PageSetup
    .CenterHeader = "&B&20" & spath & " " & reportType & "  Testing for " & todayDate
    .RightFooter = "Page: " & "&P"
    .PrintArea = "$A$2:$G$" & CStr(last_row)
    .LeftFooter = "Exported at:" & format(Now, "mm-dd-yy hh:mm")
End With


    
On Error GoTo pdf_error:
sht.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    filename:=file_save_location, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    
    
done:
Exit Sub

pdf_error:
MsgBox "PDF is unable to be generated"
Exit Sub


End Sub

