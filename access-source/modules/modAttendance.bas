Attribute VB_Name = "modAttendance"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modAttendance
' Purpose: Core business logic for Clock-In, Clock-Out, status lookup, and
'          Excel workbook persistence.
'
' Note on Access + Excel linked table constraints:
'   Microsoft Access linked Excel tables (Excel 12.0 Xml) are read-only for
'   DAO UPDATE/DELETE operations due to ACE Excel ISAM restrictions (KB 904953).
'   Therefore:
'     - Reads use DAO / Domain Aggregates against linked tables
'       (Employees, AttendanceRecords) and saved queries.
'     - Writes (INSERT for Clock-In, UPDATE for Clock-Out) use late-bound
'       ADODB via Microsoft.ACE.OLEDB.12.0 with READONLY=FALSE targeting
'       data\AttendanceRecords.xlsx ([AttendanceRecords$]).
' =============================================================================

Public Const STATUS_WORKING As String = "Working"
Public Const STATUS_COMPLETED As String = "Completed"

Private Const EXCEL_ATTENDANCE_REL_PATH As String = "data\AttendanceRecords.xlsx"
Private Const EXCEL_ATTENDANCE_RANGE As String = "[AttendanceRecords$]"
Private Const ADODB_CMD_TEXT As Long = 1
Private Const ADODB_PARAM_INPUT As Long = 1
Private Const ADODB_TYPE_VARWCHAR As Long = 202

' Resolve the physical path of AttendanceRecords.xlsx from the linked table
' connection string, falling back to CurrentProject.Path\data\AttendanceRecords.xlsx.
Public Function GetAttendanceExcelPath() As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim connectStr As String
    Dim dbPos As Long
    Dim semiPos As Long

    On Error GoTo FallbackPath
    Set db = CurrentDb
    Set tdf = db.TableDefs("AttendanceRecords")
    connectStr = tdf.Connect

    dbPos = InStr(1, connectStr, "DATABASE=", vbTextCompare)
    If dbPos > 0 Then
        dbPos = dbPos + Len("DATABASE=")
        semiPos = InStr(dbPos, connectStr, ";")
        If semiPos > 0 Then
            GetAttendanceExcelPath = Mid$(connectStr, dbPos, semiPos - dbPos)
        Else
            GetAttendanceExcelPath = Mid$(connectStr, dbPos)
        End If
        If Len(Trim$(GetAttendanceExcelPath)) > 0 Then
            Set tdf = Nothing
            Set db = Nothing
            Exit Function
        End If
    End If

FallbackPath:
    Set tdf = Nothing
    Set db = Nothing
    GetAttendanceExcelPath = CurrentProject.Path & "\" & EXCEL_ATTENDANCE_REL_PATH
End Function

' Open a writable ADODB connection to AttendanceRecords.xlsx.
Private Function OpenAttendanceExcelConnection() As Object
    Dim conn As Object
    Dim excelPath As String
    Dim connStr As String

    excelPath = GetAttendanceExcelPath()
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & excelPath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=YES;READONLY=FALSE;"";"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open connStr
    Set OpenAttendanceExcelConnection = conn
End Function

' Generate the next sequential AttendanceID in format ATT-YYYYMMDD-NNN.
Public Function GenerateAttendanceID(ByVal workDate As String) As String
    Dim dateDigits As String
    Dim prefix As String
    Dim maxId As Variant
    Dim nextSeq As Long

    dateDigits = Replace(workDate, "-", "")
    prefix = "ATT-" & dateDigits & "-"

    maxId = DMax("AttendanceID", "AttendanceRecords", "AttendanceID Like '" & SqlEscape(prefix) & "*'")

    If IsNull(maxId) Or Len(Trim$(Nz(maxId, ""))) = 0 Then
        nextSeq = 1
    Else
        nextSeq = CLng(Right$(CStr(maxId), 3)) + 1
    End If

    GenerateAttendanceID = prefix & Format$(nextSeq, "000")
End Function

' Find the AttendanceID of an employee's open ("Working") record.
' Prefers a record matching workDate when provided; otherwise returns the most
' recent open record for the employee. Returns "" if none exists.
Public Function FindOpenAttendanceID(ByVal employeeId As String, Optional ByVal workDate As String = "") As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim safeEmpId As String

    On Error GoTo Cleanup
    safeEmpId = SqlEscape(Trim$(employeeId))

    Set db = CurrentDb

    ' First check for same-day open record if workDate is supplied.
    If Len(Trim$(workDate)) > 0 Then
        sql = "SELECT TOP 1 AttendanceID FROM AttendanceRecords " & _
              "WHERE EmployeeID = '" & safeEmpId & "' " & _
              "AND WorkDate = '" & SqlEscape(Trim$(workDate)) & "' " & _
              "AND Status = '" & STATUS_WORKING & "' " & _
              "AND (ClockOut IS NULL OR ClockOut = '') " & _
              "ORDER BY ClockIn DESC, AttendanceID DESC;"
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        If Not rs.EOF Then
            FindOpenAttendanceID = CStr(rs!AttendanceID)
            GoTo Cleanup
        End If
        rs.Close
        Set rs = Nothing
    End If

    ' Fallback: latest open record across any date.
    sql = "SELECT TOP 1 AttendanceID FROM AttendanceRecords " & _
          "WHERE EmployeeID = '" & safeEmpId & "' " & _
          "AND Status = '" & STATUS_WORKING & "' " & _
          "AND (ClockOut IS NULL OR ClockOut = '') " & _
          "ORDER BY WorkDate DESC, ClockIn DESC, AttendanceID DESC;"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        FindOpenAttendanceID = CStr(rs!AttendanceID)
    Else
        FindOpenAttendanceID = ""
    End If

Cleanup:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function

' Return a Japanese display label representing the employee's current status.
Public Function GetEmployeeCurrentStatus(ByVal employeeId As String, Optional ByVal workDate As String = "") As String
    Dim targetDate As String
    Dim completedCount As Long

    If Len(Trim$(employeeId)) = 0 Then
        GetEmployeeCurrentStatus = "-"
        Exit Function
    End If

    If Len(Trim$(workDate)) = 0 Then
        targetDate = Format$(Date, "yyyy-mm-dd")
    Else
        targetDate = Trim$(workDate)
    End If

    If HasOpenAttendance(employeeId) Then
        GetEmployeeCurrentStatus = "出勤中（未退勤）"
        Exit Function
    End If

    completedCount = DCount("AttendanceID", "AttendanceRecords", _
        "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "' " & _
        "AND WorkDate = '" & SqlEscape(targetDate) & "' " & _
        "AND Status = '" & STATUS_COMPLETED & "'")

    If completedCount > 0 Then
        GetEmployeeCurrentStatus = "退勤済"
    Else
        GetEmployeeCurrentStatus = "未出勤"
    End If
End Function

' Insert a new attendance record into AttendanceRecords.xlsx via ADODB.
Private Sub InsertAttendanceRecord( _
    ByVal attendanceId As String, _
    ByVal employeeId As String, _
    ByVal workDate As String, _
    ByVal clockIn As String, _
    ByVal clockOut As String, _
    ByVal statusValue As String)

    Dim conn As Object
    Dim cmd As Object

    On Error GoTo ErrHandler
    Set conn = OpenAttendanceExcelConnection()
    Set cmd = CreateObject("ADODB.Command")

    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "INSERT INTO " & EXCEL_ATTENDANCE_RANGE & _
                      " (AttendanceID, EmployeeID, WorkDate, ClockIn, ClockOut, Status)" & _
                      " VALUES (?, ?, ?, ?, ?, ?)"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, employeeId)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, workDate)
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockIn)
    cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockOut)
    cmd.Parameters.Append cmd.CreateParameter("p6", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, statusValue)

    cmd.Execute

    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Err.Raise errNum, "modAttendance.InsertAttendanceRecord", errDesc
End Sub

' Update ClockOut and Status for an existing attendance record in AttendanceRecords.xlsx via ADODB.
Private Sub UpdateAttendanceRecordClockOut( _
    ByVal attendanceId As String, _
    ByVal clockOut As String, _
    ByVal statusValue As String)

    Dim conn As Object
    Dim cmd As Object
    Dim recordsAffected As Long

    On Error GoTo ErrHandler
    Set conn = OpenAttendanceExcelConnection()
    Set cmd = CreateObject("ADODB.Command")

    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "UPDATE " & EXCEL_ATTENDANCE_RANGE & _
                      " SET ClockOut = ?, Status = ?" & _
                      " WHERE AttendanceID = ?"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockOut)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, statusValue)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)

    cmd.Execute recordsAffected

    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    If recordsAffected = 0 Then
        Err.Raise vbObjectError + 1001, "modAttendance.UpdateAttendanceRecordClockOut", _
                  "対象の勤怠レコードが更新されませんでした (" & attendanceId & ")。"
    End If
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Err.Raise errNum, "modAttendance.UpdateAttendanceRecordClockOut", errDesc
End Sub

' Execute Clock-In workflow:
'   1. Capture current timestamp
'   2. Validate employee selection, active status, and no duplicate open record
'   3. Generate AttendanceID and append record with Status = "Working"
'   4. Return completion message
Public Function ClockInEmployee( _
    ByVal employeeId As String, _
    ByRef resultMessage As String, _
    Optional ByRef createdAttendanceId As String = "") As Boolean

    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim clockInText As String
    Dim newAttendanceId As String

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    clockInText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateClockIn(normalizedId, workDate, resultMessage) Then
        ClockInEmployee = False
        Exit Function
    End If

    newAttendanceId = GenerateAttendanceID(workDate)

    InsertAttendanceRecord _
        attendanceId:=newAttendanceId, _
        employeeId:=normalizedId, _
        workDate:=workDate, _
        clockIn:=clockInText, _
        clockOut:="", _
        statusValue:=STATUS_WORKING

    createdAttendanceId = newAttendanceId
    resultMessage = "出勤打刻が完了しました。" & vbCrLf & _
                    "従業員ID: " & normalizedId & vbCrLf & _
                    "勤怠ID: " & newAttendanceId & vbCrLf & _
                    "出勤時刻: " & clockInText
    ClockInEmployee = True
    Exit Function

ErrHandler:
    resultMessage = "出勤打刻中にエラーが発生しました: " & Err.Description
    ClockInEmployee = False
End Function

' Execute Clock-Out workflow:
'   1. Capture current timestamp
'   2. Validate employee selection, active status, and existence of open record
'   3. Find open AttendanceID and update ClockOut and Status = "Completed"
'   4. Return completion message
Public Function ClockOutEmployee( _
    ByVal employeeId As String, _
    ByRef resultMessage As String, _
    Optional ByRef updatedAttendanceId As String = "") As Boolean

    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim clockOutText As String
    Dim openAttendanceId As String

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    clockOutText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateClockOut(normalizedId, resultMessage) Then
        ClockOutEmployee = False
        Exit Function
    End If

    openAttendanceId = FindOpenAttendanceID(normalizedId, workDate)
    If Len(openAttendanceId) = 0 Then
        resultMessage = "未退勤の出勤記録が見つかりません。先に出勤打刻を行ってください。"
        ClockOutEmployee = False
        Exit Function
    End If

    UpdateAttendanceRecordClockOut _
        attendanceId:=openAttendanceId, _
        clockOut:=clockOutText, _
        statusValue:=STATUS_COMPLETED

    updatedAttendanceId = openAttendanceId
    resultMessage = "退勤打刻が完了しました。" & vbCrLf & _
                    "従業員ID: " & normalizedId & vbCrLf & _
                    "勤怠ID: " & openAttendanceId & vbCrLf & _
                    "退勤時刻: " & clockOutText
    ClockOutEmployee = True
    Exit Function

ErrHandler:
    resultMessage = "退勤打刻中にエラーが発生しました: " & Err.Description
    ClockOutEmployee = False
End Function
