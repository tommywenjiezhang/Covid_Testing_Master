Attribute VB_Name = "refreshRoster"
Sub refreshRoster()
Dim db As New testDb
Dim names As Variant
Dim util As New testUtil



names = db.getEmpName()
Dim i As Long
Dim j As Long
Dim start As Long

start = 2
If Not util.isArrayEmpty(names) Then
    empList.Unprotect
    For i = LBound(names, 2) To UBound(names, 2)
        empList.Cells(start + i, 1).value = names(0, i)
        empList.Cells(start + i, 2).value = names(1, i)
    
    Next i
    empList.Protect
End If



End Sub


Sub importBirthday()
    Dim db As New testDb
    Dim birthday As Variant
    Dim util As New testUtil
    
    birthday = db.getEmpBirthday()
    Dim i As Long
    Dim start As Long
    empBirthday.Cells.ClearContents
    
    start = 1
    If Not util.isArrayEmpty(birthday) Then
        
        For i = LBound(birthday, 2) To UBound(birthday, 2)
            With empBirthday
                .Cells(start + i, 1).value = birthday(0, i)
                .Cells(start + i, 2).value = birthday(1, i)
                
            End With
        Next i
    
    End If
    
    
End Sub


Sub importVaccine()
    Dim db As New testDb
    Dim vaccines As Variant
    Dim util As New testUtil
    Dim lastRow As Long
    
    vaccines = db.getVaccinated()
    Dim i As Long
    Dim start As Long
    empVaccine.Cells.ClearContents
    
    start = 1
    
    If Not util.isArrayEmpty(vaccines) Then
    
        For i = LBound(vaccines, 2) To UBound(vaccines, 2)
            With empVaccine
                .Cells(start + i, 1).value = vaccines(0, i)
                .Cells(start + i, 2).value = vaccines(1, i)
                
            End With
        Next i
        
    With empVaccine
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range("A1:B" & lastRow).RemoveDuplicates columns:=Array(1, 2)
    End With
    
    End If
    
    
End Sub

Sub addVaccine()
    Dim db As New testDb
    Dim empName As String
    Dim util As New testUtil
    
        
    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, empList.Range("B2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
            empName = ActiveCell.value
            db.insertVaccine (empName)
            MsgBox "Employee successfully add to vaccination list"
        End If
    End If
    
    
    
    
End Sub



Sub lookupVaccine()
    
    Dim last_row  As Long
    Dim vaccine_rng As Range
    

    With empList
        .Unprotect
        last_row = .Cells(.Rows.count, "A").End(xlUp).Row
        
        .Range("E2:E" & last_row).Interior.ColorIndex = 0
        
        Range("E2:E" & last_row).FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(TRIM([@[Employee ID]]),empVaccine!R1C1:R700C2,1,FALSE)),""No Vaccine"",""vaccinated"")"
        
        
        For Each c In .Range("E2:E" & last_row)
            
            If c.value = "vaccinated" Then
                c.Interior.color = RGB(124, 252, 0)
            End If
            
            
        Next c
        .Protect
    End With
    
End Sub




