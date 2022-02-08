VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pivotReport 
   Caption         =   "Testing Summary Report"
   ClientHeight    =   5172
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   6750
   OleObjectBlob   =   "pivotReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pivotReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnClose_Click()
 Unload Me
End Sub

Private Sub btnSumit_Click()
    Dim startDate  As String
    Dim endDate As String
    
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
    Dim Result As Variant
    Dim db As New testDb
    Dim startDate As Date
    Dim endDate As Date
    Dim util As New testUtil
    Dim new_wb As Workbook
    Dim filename As String
    
    Dim data_sht As Worksheet
   
    If Not IsError(CDate(startDateStr)) And Not IsError(CDate(endDateStr)) Then
        startDate = CDate(startDateStr)
        endDate = CDate(endDateStr)
        
        Dim i As Long
        Dim j As Long
    
        Dim start As Long
        start = 2
        
        Result = db.getTestHistory(startDate, endDate, True, True)
        If Not util.isArrayEmpty(Result) Then
            Set new_wb = Workbooks.Add
            Set data_sht = new_wb.Sheets(1)
            filename = "Weekly Report for " & format(startDate, "mm-dd-yy")
            With data_sht
                .Cells(1, 1).value = "empName"
                .Cells(1, 2).value = "TestDate"
                .Cells(1, 3).value = "typeOfTest"
                .Cells(1, 4).value = "Category"
            End With
            
            For j = LBound(Result, 2) To UBound(Result, 2)
                With data_sht
                    .Cells(start + j, 1).value = Result(3, j)
                    .Cells(start + j, 2).value = Result(1, j)
                    .Cells(start + j, 3).value = Result(2, j)
                    .Cells(start + j, 4).value = Result(4, j)
                End With
                 
            Next j
            createPivotTable new_wb
            weekTotalPivot new_wb
            On Error GoTo report_not_save
                new_wb.SaveAs filename:=ThisWorkbook.path & "\" & filename & ".xlsx"
            
            If Me.generatePdfBtn.value = True Then
                Shell "taskkill /IM ""AcroRd32.exe"" /F"
                generatePdf new_wb, startDate, filename & ".pdf"
            End If
            
            If Me.sendEmailCopyBtn.value = True Then
                sendEmail new_wb, filename & ".xlsx", startDate
            End If
            
            Unload Me
            
        End If
    Else
        Me.endDateTxt.BackColor = RGB(255, 255, 0)
        Me.startDateTxt.BackColor = RGB(255, 255, 0)
        Me.warning.Visible = True
    End If

done:
    Exit Sub
report_not_save:
    MsgBox "Please close any existing report and try again"
End Sub

Private Function sendEmail(ByRef wb As Workbook, ByVal filename As String, ByVal startDate As Date)
    Dim EmailApp As Outlook.Application
    Set EmailApp = New Outlook.Application
    
    Dim EmailItem As Outlook.MailItem
    Set EmailItem = EmailApp.CreateItem(olMailItem)
    
    
    
    With EmailItem
        .Subject = "Weekly Report for " & format(startDate, "mm-dd-yy")
        .HTMLBody = "<h1>Attached Weekly Report for " & format(startDate, "mm-dd-yy") & "</h1>"
        .Attachments.Add ThisWorkbook.path & "\" & filename
        .Display
    End With
    
    Set EmailItem = Nothing
    Set EmailApp = Nothing
End Function
Private Function generatePdf(ByRef wb As Workbook, ByVal startDate As Date, ByVal filepath As String) As String
    Dim tfo As New TestExport
    Dim new_sht As Worksheet
    
    With wb
        Set new_sht = .Sheets(Sheets.count)
    End With
    new_sht.UsedRange.columns.AutoFit

    With new_sht.PageSetup
        .CenterHeader = "&B&20" & "Weekly Report for " & format(startDate, "mm-dd-yy")
        .RightFooter = "Page: " & "&P"
        .CenterHorizontally = True
        .PrintArea = new_sht.PivotTables(1).TableRange2.Address
    End With

    On Error GoTo pdf_error
    new_sht.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        filename:=ThisWorkbook.path & "\" & filepath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    generatePdf = filepath

done:
    Exit Function
    
pdf_error:
    MsgBox "Cannot create PDF report"
    
End Function
Private Sub createPivotTable(ByRef wb As Workbook)
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    Dim pvt_sht As Worksheet
    Dim data_rng As Range
    
    With wb
        .Sheets(1).name = "Data"
        Set data_rng = .Sheets(1).UsedRange
        Set pvtCache = .PivotCaches.Create _
        (SourceType:=xlDatabase, SourceData:=data_rng)
        Set pvt_sht = .Sheets.Add
    End With
    
    
    Set pvt = pvtCache.createPivotTable(TableDestination:=pvt_sht.Cells(1, 1), TableName:="TestingReport")
    
    With pvt.PivotFields("empName")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    
    With pvt.PivotFields("TestDate")
        .Orientation = xlRowField
        .DataRange.Cells(2).Group start:=True, End:=True, BY:=7, _
        Periods:=Array(False, False, False, True, False, False, False)
        .Position = 2
    End With
    
    
    With pvt.PivotFields("typeOfTest")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pvt.PivotFields("typeOfTest")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlCount
    End With
    
    
    'pvt.PivotFields("empName").Subtotals(1) = False
    
    With pvt
        For Each c In .PivotFields("typeOfTest").PivotItems("RAPID").DataRange.Cells
          If c.value = "" Or CInt(c.value) <= 1 Then
            With c
              .Style = "Bad"
            End With
          End If
        Next
    End With
    
    With pvt_sht
        .Range("B1").value = "Test Total"
        .Range("A1").value = "Weekly Summary"
        .Range("A1:D1").columns.AutoFit
    End With
    
    pvt_sht.name = "Test Weekly Summary"
    
End Sub
Private Sub createPivotChart(ByRef wb As Workbook)
    Dim ch As Chart
    Dim rng As Range
    
    With wb
        Set ch = .Charts.Add(After:=.Sheets(.Sheets.count))
        Set rng = .Sheets("Weekly Total").PivotTables(1).TableRange2
    End With
    
    With ch
        .SetSourceData Source:=rng
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "Weekly Total"
        .name = "Total by Test Chart"
        .SeriesCollection(1). _
        ApplyDataLabels Type:=xlDataLabelsShowValue, ShowValue:=True
        .SeriesCollection(2). _
        ApplyDataLabels Type:=xlDataLabelsShowValue, ShowValue:=True
    End With
    
End Sub

Private Sub weekTotalPivot(ByRef wb As Workbook)
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    Dim pvt_sht As Worksheet
    Dim data_rng As Range
    
    With wb
        .Sheets(1).name = "Weeks Total"
        Set data_rng = .Worksheets("Data").UsedRange
        Set pvtCache = .PivotCaches.Create _
        (SourceType:=xlDatabase, SourceData:=data_rng)
        Set pvt_sht = .Sheets.Add(After:=.Sheets(.Sheets.count))
    End With
    
    
    Set pvt = pvtCache.createPivotTable(TableDestination:=pvt_sht.Cells(1, 1), TableName:="TestingReport")
    
    
    With pvt.PivotFields("Category")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    
    With pvt.PivotFields("TestDate")
        .Orientation = xlRowField
        .DataRange.Cells(2).Group _
        Periods:=Array(False, False, False, True, False, False, False)
        .Position = 2
    End With
    
      With pvt.PivotFields("typeOfTest")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pvt.PivotFields("typeOfTest")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlCount
    End With
    
    With pvt_sht
        .Range("B1").value = "Test Total"
        .Range("A1").value = "Weekly Total"
        .Range("A1:D1").columns.AutoFit
    End With
    
    pvt_sht.name = "Weekly Total"
    createPivotChart wb
End Sub




Private Sub UserForm_Initialize()
    Me.endDateTxt.value = format(Date, "mm/dd/yyyy")
End Sub
