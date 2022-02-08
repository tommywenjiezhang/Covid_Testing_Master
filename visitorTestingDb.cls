VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "visitorTestingDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private dbPath As String
Private TableName As String

Private Sub Class_Initialize()
    dbPath = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & Environ("USERPROFILE") & "\Desktop" & "\Testingdb.accdb;"
                    
End Sub



Public Function insertTesting(name As String, timeIn As Date, symptom As Boolean, typeOfTest As String, dob As Date)
    Set conn = New ADODB.Connection
    
    With conn
    .Open dbPath
    .Execute "INSERT INTO visitorTesting" & _
        " ( visitorName, timeTested, symptom, typeOfTest,visitorDOB)" & _
        " VALUES(" & _
        " '" & name & "'," & _
        format(timeIn, "\#yyyy-mm-dd hh:nn:ss\#") & "," & _
        CStr(symptom) & "," & _
        "'" & typeOfTest & "', " & _
        format(dob, "\#yyyy-mm-dd hh:nn:ss\#") & ")"
    .Close
    End With
    
    Set conn = Nothing
    
End Function


Public Function getTestHistory(startDate As Date, endDate As Date, Optional agg As Boolean = False) As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim name As String
    Dim util As New testUtil
    Dim checkInDate As String
    Dim output As Variant
    Dim Result As Variant
    
    Dim query As String
    
    query = "Select ID, timeTested, typeOfTest,visitorName from visitorTesting where " & _
    "timeTested >= " & format(startDate, "\#yyyy-mm-dd\#") & " and timeTested < " & format(endDate, "\#yyyy-mm-dd\#") & _
    " order by timeTested desc"
    
    If agg Then
    
    query = "SELECT empList.empID, [Testing].[timeTested] AS timeTested, Testing.typeOfTest AS typeOfTest, empList.empName" & _
    " FROM empList " & _
    " LEFT JOIN Testing ON empList.empID = Testing.empID " & _
    " where Testing.timeTested <= " & format(endDate, "\#yyyy-mm-dd\#") & _
    " and Testing.timeTested >= " & format(startDate, "\#yyyy-mm-dd\#")

    End If
    
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    
    
    If Not rs.BOF Or Not rs.EOF Then
        rs.MoveFirst
        Result = rs.GetRows
    End If
    
    
    
    
    rs.Close
    conn.Close
    getTestHistory = Result
    
    Set rs = Nothing
    Set conn = Nothing
    
End Function
