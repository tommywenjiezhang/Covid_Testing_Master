Attribute VB_Name = "exportPDF"
Function save_as_pdf(spath As String)

Dim tfo As New TestExport

Dim todayDate As String
Dim filepath As String
Dim last_row As Long
Dim vist_last_row As Long
Dim emp_save_location As String, vist_save_location As String


With testRoster
    last_row = .Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Row
    .Cells.EntireColumn.AutoFit
    .Range("G3:G" & last_row).Interior.ColorIndex = 0
    
End With
With visitorTesting
    vist_last_row = .Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Row
    .Cells.EntireColumn.AutoFit
End With


todayDate = format(Now, "dddd dd mmm, yyyy")
filepath = format(Now, "mm-dd-yy")


emp_save_location = tfo.full_path & "\pdf\" & filepath & spath & "emp-screening.pdf"
vist_save_location = tfo.full_path & "\pdf\" & filepath & spath & "vistor-screening.pdf"




With testRoster.PageSetup
    .CenterHeader = "&B&20" & spath & " Employee Testing for " & todayDate
    .RightFooter = "Page: " & "&P"
    .PrintArea = "$A$2:$G$" & CStr(last_row)
    
End With

With visitorTesting.PageSetup
    .CenterHeader = "&B&20" & spath & " Visitor Testing for " & todayDate
    .RightFooter = "Page: " & "&P"
    .PrintArea = "$A$2:$F$" & CStr(vist_last_row)
    
End With

    
On Error GoTo pdf_error
testRoster.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    filename:=emp_save_location, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    
visitorTesting.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    filename:=vist_save_location, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=False, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True

done:
Exit Function

pdf_error:
MsgBox "PDF is unable to be generated"
Exit Function


End Function
