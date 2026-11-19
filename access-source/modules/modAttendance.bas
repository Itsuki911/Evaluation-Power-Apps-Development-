Attribute VB_Name = "modAttendance"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modAttendance
' Purpose: Core business logic for Clock-In, Break Start/End, Clock-Out,
'          Correction Request submission/approval (with closed-month override),
'          and Monthly Closing lock operations.
' History:
'   2011/04/01 山本: 初期出退勤処理実装
'   2015/05/12 佐藤: 休憩開始・終了打刻と自動休憩算出を追加
'   2017/10/04 鈴木: 打刻修正申請・上長承認フローを追加
'   2018/11/20 高橋: 月次締め管理・ADODB Excel書込対応 (KB 904953対策)
' =============================================================================

Private Const EXCEL_ATT_REL_PATH As String = "data\AttendanceRecords.xlsx"
Private Const EXCEL_ATT_RANGE As String = "[AttendanceRecords$]"
Private Const EXCEL_REQ_REL_PATH As String = "data\CorrectionRequests.xlsx"
Private Const EXCEL_REQ_RANGE As String = "[CorrectionRequests$]"
Private Const EXCEL_CLS_REL_PATH As String = "data\MonthlyClosings.xlsx"
Private Const EXCEL_CLS_RANGE As String = "[MonthlyClosings$]"

' Generate sequential ID: e.g. ATT-20261120-001 or REQ-20261120-001
Public Function GenerateNextID( _
    ByVal prefixTag As String, _
    ByVal tableName As String, _
    ByVal idColumn As String, _
    ByVal workDate As String) As String

    Dim dateDigits As String
    Dim prefix As String
    Dim maxId As Variant
    Dim nextSeq As Long

    dateDigits = Replace(workDate, "-", "")
    prefix = prefixTag & "-" & dateDigits & "-"
    maxId = DMax(idColumn, tableName, idColumn & " Like '" & SqlEscape(prefix) & "*'")

    If IsNull(maxId) Or Len(Trim$(Nz(maxId, ""))) = 0 Then
        nextSeq = 1
    Else
        nextSeq = CLng(Right$(CStr(maxId), 3)) + 1
    End If

    GenerateNextID = prefix & Format$(nextSeq, "000")
End Function

' Find an employee's open attendance record (Working or OnBreak).
Public Function FindOpenAttendanceID(ByVal employeeId As String, Optional ByVal workDate As String = "") As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim safeEmpId As String

    On Error GoTo Cleanup
    safeEmpId = SqlEscape(Trim$(employeeId))
    Set db = CurrentDb

    If Len(Trim$(workDate)) > 0 Then
        sql = "SELECT TOP 1 AttendanceID FROM AttendanceRecords " & _
              "WHERE EmployeeID = '" & safeEmpId & "' " & _
              "AND WorkDate = '" & SqlEscape(Trim$(workDate)) & "' " & _
              "AND Status IN ('" & ATT_STATUS_WORKING & "', '" & ATT_STATUS_ONBREAK & "') " & _
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

    sql = "SELECT TOP 1 AttendanceID FROM AttendanceRecords " & _
          "WHERE EmployeeID = '" & safeEmpId & "' " & _
          "AND Status IN ('" & ATT_STATUS_WORKING & "', '" & ATT_STATUS_ONBREAK & "') " & _
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

' Return Japanese status label for frmMain display.
Public Function GetEmployeeCurrentStatus(ByVal employeeId As String, Optional ByVal workDate As String = "") As String
    Dim targetDate As String
    Dim cleanId As String

    cleanId = Trim$(employeeId)
    If Len(cleanId) = 0 Then
        GetEmployeeCurrentStatus = "-"
        Exit Function
    End If

    If Len(Trim$(workDate)) = 0 Then
        targetDate = Format$(Date, "yyyy-mm-dd")
    Else
        targetDate = Trim$(workDate)
    End If

    If HasOnBreakAttendance(cleanId) Then
        GetEmployeeCurrentStatus = "休憩中"
        Exit Function
    End If

    If HasOpenAttendance(cleanId) Then
        GetEmployeeCurrentStatus = "出勤中（未退勤）"
        Exit Function
    End If

    If DCount("AttendanceID", "AttendanceRecords", _
        "EmployeeID = '" & SqlEscape(cleanId) & "' AND WorkDate = '" & SqlEscape(targetDate) & "' AND Status = '" & ATT_STATUS_LEAVE & "'") > 0 Then
        GetEmployeeCurrentStatus = "休暇登録済"
        Exit Function
    End If

    If DCount("AttendanceID", "AttendanceRecords", _
        "EmployeeID = '" & SqlEscape(cleanId) & "' AND WorkDate = '" & SqlEscape(targetDate) & "' AND Status = '" & ATT_STATUS_COMPLETED & "'") > 0 Then
        GetEmployeeCurrentStatus = "退勤済"
    Else
        GetEmployeeCurrentStatus = "未出勤"
    End If
End Function

' Lookup employee's scheduled start or end time ("HH:MM") from WorkPatterns.
Private Function LookupEmployeeScheduleTime(ByVal employeeId As String, ByVal fieldName As String) As String
    Dim ptnCode As Variant
    Dim timeVal As Variant

    ptnCode = DLookup("PatternCode", "Employees", "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "'")
    If IsNull(ptnCode) Then
        LookupEmployeeScheduleTime = ""
        Exit Function
    End If

    timeVal = DLookup(fieldName, "WorkPatterns", "PatternCode = '" & SqlEscape(CStr(ptnCode)) & "'")
    LookupEmployeeScheduleTime = Trim$(Nz(timeVal, ""))
End Function

' --- ADODB persistence helpers ---

Private Sub InsertAttendanceRow( _
    ByVal attendanceId As String, _
    ByVal employeeId As String, _
    ByVal workDate As String, _
    ByVal clockIn As String, _
    ByVal clockOut As String, _
    ByVal breakStart As String, _
    ByVal breakEnd As String, _
    ByVal breakMinutes As String, _
    ByVal workType As String, _
    ByVal statusValue As String, _
    ByVal remarks As String, _
    ByVal updatedBy As String)

    Dim conn As Object
    Dim cmd As Object

    On Error GoTo ErrHandler
    Set conn = OpenExcelWriteConnection("AttendanceRecords", EXCEL_ATT_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "INSERT INTO " & EXCEL_ATT_RANGE & _
                      " (AttendanceID, EmployeeID, WorkDate, ClockIn, ClockOut, BreakStart, BreakEnd, BreakMinutes, WorkType, Status, Remarks, UpdatedBy)" & _
                      " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, employeeId)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, workDate)
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockIn)
    cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockOut)
    cmd.Parameters.Append cmd.CreateParameter("p6", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, breakStart)
    cmd.Parameters.Append cmd.CreateParameter("p7", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, breakEnd)
    cmd.Parameters.Append cmd.CreateParameter("p8", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, breakMinutes)
    cmd.Parameters.Append cmd.CreateParameter("p9", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, workType)
    cmd.Parameters.Append cmd.CreateParameter("p10", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, statusValue)
    cmd.Parameters.Append cmd.CreateParameter("p11", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 255, remarks)
    cmd.Parameters.Append cmd.CreateParameter("p12", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, updatedBy)

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
    Err.Raise errNum, "modAttendance.InsertAttendanceRow", errDesc
End Sub

Private Sub UpdateAttendanceBreakStart(ByVal attendanceId As String, ByVal breakStart As String, ByVal updatedBy As String)
    Dim conn As Object
    Dim cmd As Object

    On Error GoTo ErrHandler
    Set conn = OpenExcelWriteConnection("AttendanceRecords", EXCEL_ATT_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "UPDATE " & EXCEL_ATT_RANGE & _
                      " SET BreakStart = ?, Status = ?, UpdatedBy = ?" & _
                      " WHERE AttendanceID = ?"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, breakStart)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, ATT_STATUS_ONBREAK)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, updatedBy)
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)
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
    Err.Raise errNum, "modAttendance.UpdateAttendanceBreakStart", errDesc
End Sub

Private Sub UpdateAttendanceBreakEnd(ByVal attendanceId As String, ByVal breakEnd As String, ByVal breakMinText As String, ByVal updatedBy As String)
    Dim conn As Object
    Dim cmd As Object

    On Error GoTo ErrHandler
    Set conn = OpenExcelWriteConnection("AttendanceRecords", EXCEL_ATT_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "UPDATE " & EXCEL_ATT_RANGE & _
                      " SET BreakEnd = ?, BreakMinutes = ?, Status = ?, UpdatedBy = ?" & _
                      " WHERE AttendanceID = ?"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, breakEnd)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, breakMinText)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, ATT_STATUS_WORKING)
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, updatedBy)
    cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)
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
    Err.Raise errNum, "modAttendance.UpdateAttendanceBreakEnd", errDesc
End Sub

Private Sub UpdateAttendanceClockOut( _
    ByVal attendanceId As String, _
    ByVal clockOut As String, _
    ByVal breakMinText As String, _
    ByVal workType As String, _
    ByVal statusValue As String, _
    ByVal updatedBy As String)

    Dim conn As Object
    Dim cmd As Object
    Dim affected As Long

    On Error GoTo ErrHandler
    Set conn = OpenExcelWriteConnection("AttendanceRecords", EXCEL_ATT_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "UPDATE " & EXCEL_ATT_RANGE & _
                      " SET ClockOut = ?, BreakMinutes = ?, WorkType = ?, Status = ?, UpdatedBy = ?" & _
                      " WHERE AttendanceID = ?"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, clockOut)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, breakMinText)
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, workType)
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, statusValue)
    cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, updatedBy)
    cmd.Parameters.Append cmd.CreateParameter("p6", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attendanceId)
    cmd.Execute affected

    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    If affected = 0 Then
        Err.Raise vbObjectError + 1001, "modAttendance.UpdateAttendanceClockOut", _
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
    Err.Raise errNum, "modAttendance.UpdateAttendanceClockOut", errDesc
End Sub

' --- Public business workflows ---

' 1. Clock-In workflow
Public Function ClockInEmployee( _
    ByVal employeeId As String, _
    ByRef resultMessage As String, _
    Optional ByRef createdAttendanceId As String = "", _
    Optional ByVal workTypeOverride As String = "") As Boolean

    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim clockInText As String
    Dim newId As String
    Dim schedStart As String
    Dim lateMin As Long
    Dim workTypeCode As String

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    clockInText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateClockIn(normalizedId, workDate, resultMessage) Then
        ClockInEmployee = False
        Exit Function
    End If

    If Len(Trim$(workTypeOverride)) > 0 Then
        workTypeCode = Trim$(workTypeOverride)
    Else
        schedStart = LookupEmployeeScheduleTime(normalizedId, "StartTime")
        lateMin = JudgeLateMinutes(workDate, clockInText, schedStart)
        If lateMin > 0 Then
            workTypeCode = WT_LATE_EARLY
        Else
            workTypeCode = WT_NORMAL
        End If
    End If

    newId = GenerateNextID("ATT", "AttendanceRecords", "AttendanceID", workDate)

    InsertAttendanceRow _
        attendanceId:=newId, _
        employeeId:=normalizedId, _
        workDate:=workDate, _
        clockIn:=clockInText, _
        clockOut:="", _
        breakStart:="", _
        breakEnd:="", _
        breakMinutes:="", _
        workType:=workTypeCode, _
        statusValue:=ATT_STATUS_WORKING, _
        remarks:="", _
        updatedBy:=normalizedId

    createdAttendanceId = newId
    resultMessage = "出勤打刻が完了しました。" & vbCrLf & _
                    "勤怠ID: " & newId & vbCrLf & _
                    "出勤時刻: " & clockInText & vbCrLf & _
                    "区分: " & GetWorkTypeLabel(workTypeCode)
    ClockInEmployee = True
    Exit Function

ErrHandler:
    gLastErrorMsg = Err.Description
    resultMessage = "出勤打刻中にエラーが発生しました: " & Err.Description
    ClockInEmployee = False
End Function

' 2. Break-Start workflow (Working -> OnBreak)
Public Function StartBreakEmployee(ByVal employeeId As String, ByRef resultMessage As String) As Boolean
    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim breakStartText As String
    Dim openId As String

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    breakStartText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateBreakStart(normalizedId, workDate, resultMessage) Then
        StartBreakEmployee = False
        Exit Function
    End If

    openId = FindOpenAttendanceID(normalizedId, workDate)
    If Len(openId) = 0 Then
        resultMessage = "対象の勤務レコードが見つかりません。"
        StartBreakEmployee = False
        Exit Function
    End If

    UpdateAttendanceBreakStart openId, breakStartText, normalizedId
    resultMessage = "休憩開始を記録しました。(" & breakStartText & ")"
    StartBreakEmployee = True
    Exit Function

ErrHandler:
    gLastErrorMsg = Err.Description
    resultMessage = "休憩開始打刻中にエラーが発生しました: " & Err.Description
    StartBreakEmployee = False
End Function

' 3. Break-End workflow (OnBreak -> Working)
Public Function EndBreakEmployee(ByVal employeeId As String, ByRef resultMessage As String) As Boolean
    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim breakEndText As String
    Dim openId As String
    Dim prevBreakStart As Variant
    Dim prevBreakMin As Variant
    Dim addedMin As Long
    Dim totalBreakMin As Long

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    breakEndText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateBreakEnd(normalizedId, workDate, resultMessage) Then
        EndBreakEmployee = False
        Exit Function
    End If

    openId = FindOpenAttendanceID(normalizedId, workDate)
    If Len(openId) = 0 Then
        resultMessage = "休憩中のレコードが見つかりません。"
        EndBreakEmployee = False
        Exit Function
    End If

    prevBreakStart = DLookup("BreakStart", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'")
    prevBreakMin = DLookup("BreakMinutes", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'")

    If Not IsNull(prevBreakStart) And IsDate(prevBreakStart) Then
        addedMin = DateDiff("n", CDate(prevBreakStart), punchTime)
        If addedMin < 0 Then addedMin = 0
    Else
        addedMin = 0
    End If

    If IsNumeric(Nz(prevBreakMin, "")) Then
        totalBreakMin = CLng(prevBreakMin) + addedMin
    Else
        totalBreakMin = addedMin
    End If

    UpdateAttendanceBreakEnd openId, breakEndText, CStr(totalBreakMin), normalizedId
    resultMessage = "休憩終了を記録しました。(" & breakEndText & " / 累計休憩: " & totalBreakMin & "分)"
    EndBreakEmployee = True
    Exit Function

ErrHandler:
    gLastErrorMsg = Err.Description
    resultMessage = "休憩終了打刻中にエラーが発生しました: " & Err.Description
    EndBreakEmployee = False
End Function

' 4. Clock-Out workflow (Working -> Completed)
Public Function ClockOutEmployee( _
    ByVal employeeId As String, _
    ByRef resultMessage As String, _
    Optional ByRef updatedAttendanceId As String = "") As Boolean

    Dim normalizedId As String
    Dim punchTime As Date
    Dim workDate As String
    Dim clockOutText As String
    Dim openId As String
    Dim recWorkDate As String
    Dim recClockIn As String
    Dim recBreakStart As Variant
    Dim recBreakEnd As Variant
    Dim recBreakMin As Variant
    Dim recWorkType As String
    Dim grossMin As Long
    Dim effBreakMin As Long
    Dim netWorkMin As Long
    Dim overtimeMin As Long
    Dim lateNightMin As Long
    Dim schedEnd As String

    On Error GoTo ErrHandler
    normalizedId = Trim$(employeeId)
    punchTime = Now()
    workDate = Format$(punchTime, "yyyy-mm-dd")
    clockOutText = Format$(punchTime, "yyyy-mm-dd hh:nn:ss")

    If Not ValidateClockOut(normalizedId, resultMessage) Then
        ClockOutEmployee = False
        Exit Function
    End If

    openId = FindOpenAttendanceID(normalizedId, workDate)
    If Len(openId) = 0 Then
        resultMessage = "未退勤の出勤記録が見つかりません。先に出勤打刻を行ってください。"
        ClockOutEmployee = False
        Exit Function
    End If

    recWorkDate = CStr(Nz(DLookup("WorkDate", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'"), workDate))
    If IsMonthClosed(Left$(recWorkDate, 7)) Then
        resultMessage = "対象月 (" & Left$(recWorkDate, 7) & ") は締済のため直接退勤できません。修正申請を利用してください。"
        ClockOutEmployee = False
        Exit Function
    End If

    recClockIn = CStr(Nz(DLookup("ClockIn", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'"), ""))
    recBreakStart = DLookup("BreakStart", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'")
    recBreakEnd = DLookup("BreakEnd", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'")
    recBreakMin = DLookup("BreakMinutes", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'")
    recWorkType = CStr(Nz(DLookup("WorkType", "AttendanceRecords", "AttendanceID = '" & SqlEscape(openId) & "'"), WT_NORMAL))

    grossMin = RoundClockOutMinutes(recWorkDate, clockOutText, 15) - RoundClockInMinutes(recWorkDate, recClockIn, 15)
    If grossMin < 0 Then grossMin = 0
    effBreakMin = CalcEffectiveBreakMinutes(grossMin, recBreakStart, recBreakEnd, recBreakMin)
    netWorkMin = CalcWorkMinutes(recWorkDate, recClockIn, clockOutText, recBreakStart, recBreakEnd, CStr(effBreakMin), 15)
    overtimeMin = CalcOvertimeMinutes(netWorkMin, 480)
    lateNightMin = CalcLateNightMinutes(recWorkDate, recClockIn, clockOutText, 15)

    ' Flag early leave if WorkType is normal and ClockOut precedes pattern EndTime
    If recWorkType = WT_NORMAL Then
        schedEnd = LookupEmployeeScheduleTime(normalizedId, "EndTime")
        If JudgeEarlyLeaveMinutes(recWorkDate, clockOutText, schedEnd) > 0 Then
            recWorkType = WT_LATE_EARLY
        End If
    End If

    UpdateAttendanceClockOut _
        attendanceId:=openId, _
        clockOut:=clockOutText, _
        breakMinText:=CStr(effBreakMin), _
        workType:=recWorkType, _
        statusValue:=ATT_STATUS_COMPLETED, _
        updatedBy:=normalizedId

    updatedAttendanceId = openId
    resultMessage = "退勤打刻が完了しました。" & vbCrLf & _
                    "勤怠ID: " & openId & vbCrLf & _
                    "退勤時刻: " & clockOutText & vbCrLf & _
                    "実働: " & FormatMinutesHHMM(netWorkMin) & _
                    " / 残業: " & FormatMinutesHHMM(overtimeMin) & _
                    " / 深夜: " & FormatMinutesHHMM(lateNightMin)
    ClockOutEmployee = True
    Exit Function

ErrHandler:
    gLastErrorMsg = Err.Description
    resultMessage = "退勤打刻中にエラーが発生しました: " & Err.Description
    ClockOutEmployee = False
End Function

' 5. Submit correction request (CorrectionRequests.xlsx)
Public Function SubmitCorrectionRequest( _
    ByVal attendanceId As String, _
    ByVal employeeId As String, _
    ByVal targetDate As String, _
    ByVal reqClockIn As String, _
    ByVal reqClockOut As String, _
    ByVal reqBreakMin As String, _
    ByVal reqWorkType As String, _
    ByVal reasonText As String, _
    ByRef resultMessage As String, _
    Optional ByRef createdReqId As String = "") As Boolean

    Dim conn As Object
    Dim cmd As Object
    Dim newReqId As String
    Dim todayStr As String

    On Error GoTo ErrHandler
    If Not ValidateCorrectionInput(employeeId, targetDate, reqClockIn, reqClockOut, reqBreakMin, reqWorkType, reasonText, resultMessage) Then
        SubmitCorrectionRequest = False
        Exit Function
    End If

    todayStr = Format$(Date, "yyyy-mm-dd")
    newReqId = GenerateNextID("REQ", "CorrectionRequests", "RequestID", todayStr)

    Set conn = OpenExcelWriteConnection("CorrectionRequests", EXCEL_REQ_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT
    cmd.CommandText = "INSERT INTO " & EXCEL_REQ_RANGE & _
                      " (RequestID, AttendanceID, EmployeeID, TargetDate, ReqClockIn, ReqClockOut, ReqBreakMin, ReqWorkType, Reason, ApprovalStatus, ApproverID, ApprovedAt)" & _
                      " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, newReqId)
    cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, Trim$(attendanceId))
    cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, Trim$(employeeId))
    cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, Trim$(targetDate))
    cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, Trim$(reqClockIn))
    cmd.Parameters.Append cmd.CreateParameter("p6", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, Trim$(reqClockOut))
    cmd.Parameters.Append cmd.CreateParameter("p7", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, Trim$(reqBreakMin))
    cmd.Parameters.Append cmd.CreateParameter("p8", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, Trim$(reqWorkType))
    cmd.Parameters.Append cmd.CreateParameter("p9", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 255, Trim$(reasonText))
    cmd.Parameters.Append cmd.CreateParameter("p10", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, APPR_PENDING)
    cmd.Parameters.Append cmd.CreateParameter("p11", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, "")
    cmd.Parameters.Append cmd.CreateParameter("p12", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, "")
    cmd.Execute

    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    createdReqId = newReqId
    resultMessage = "打刻修正申請を登録しました。(申請ID: " & newReqId & ")"
    SubmitCorrectionRequest = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    resultMessage = "修正申請登録エラー: " & Err.Description
    SubmitCorrectionRequest = False
End Function

' 6. Manager/Admin approval or rejection of a pending correction request.
'    When approved ("1"), applies changes to AttendanceRecords even if the target month is closed.
Public Function ApproveOrRejectCorrection( _
    ByVal requestId As String, _
    ByVal approverId As String, _
    ByVal newStatus As String, _
    ByRef resultMessage As String) As Boolean

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim connReq As Object
    Dim cmdReq As Object
    Dim connAtt As Object
    Dim cmdAtt As Object
    Dim nowText As String
    Dim attId As String
    Dim empId As String
    Dim targetDate As String
    Dim reqIn As String
    Dim reqOut As String
    Dim reqBreak As String
    Dim reqWT As String
    Dim attStatus As String

    On Error GoTo ErrHandler
    If Not IsManagerOrAdmin(approverId) Then
        resultMessage = "承認・差戻権限がありません (Manager または Admin 権限が必要です)。"
        ApproveOrRejectCorrection = False
        Exit Function
    End If

    If newStatus <> APPR_APPROVED And newStatus <> APPR_REJECTED Then
        resultMessage = "不正な承認ステータスです (" & newStatus & ")。"
        ApproveOrRejectCorrection = False
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM CorrectionRequests WHERE RequestID = '" & SqlEscape(Trim$(requestId)) & "';", _
        dbOpenSnapshot)

    If rs.EOF Then
        resultMessage = "対象の修正申請が見つかりません (" & requestId & ")。"
        GoTo CleanupFail
    End If

    If Trim$(Nz(rs!ApprovalStatus, "")) <> APPR_PENDING Then
        resultMessage = "この申請は既に処理済です (" & GetApprovalStatusLabel(rs!ApprovalStatus) & ")。"
        GoTo CleanupFail
    End If

    attId = Trim$(Nz(rs!AttendanceID, ""))
    empId = Trim$(Nz(rs!EmployeeID, ""))
    targetDate = Trim$(Nz(rs!TargetDate, ""))
    reqIn = Trim$(Nz(rs!ReqClockIn, ""))
    reqOut = Trim$(Nz(rs!ReqClockOut, ""))
    reqBreak = Trim$(Nz(rs!ReqBreakMin, ""))
    reqWT = Trim$(Nz(rs!ReqWorkType, WT_NORMAL))
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    nowText = Format$(Now(), "yyyy-mm-dd hh:nn:ss")

    ' If Approved, reflect correction into AttendanceRecords (bypasses monthly close lock!)
    If newStatus = APPR_APPROVED Then
        If reqWT = WT_PAID_LEAVE Or reqWT = WT_ABSENT Then
            attStatus = ATT_STATUS_LEAVE
        Else
            attStatus = ATT_STATUS_COMPLETED
        End If

        If Len(attId) > 0 And DCount("AttendanceID", "AttendanceRecords", "AttendanceID = '" & SqlEscape(attId) & "'") > 0 Then
            Set connAtt = OpenExcelWriteConnection("AttendanceRecords", EXCEL_ATT_REL_PATH)
            Set cmdAtt = CreateObject("ADODB.Command")
            Set cmdAtt.ActiveConnection = connAtt
            cmdAtt.CommandType = ADODB_CMD_TEXT
            cmdAtt.CommandText = "UPDATE " & EXCEL_ATT_RANGE & _
                                 " SET ClockIn = ?, ClockOut = ?, BreakMinutes = ?, WorkType = ?, Status = ?, Remarks = ?, UpdatedBy = ?" & _
                                 " WHERE AttendanceID = ?"
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, reqIn)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, reqOut)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, reqBreak)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, reqWT)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attStatus)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p6", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 255, requestId & "承認反映")
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p7", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, approverId)
            cmdAtt.Parameters.Append cmdAtt.CreateParameter("p8", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, attId)
            cmdAtt.Execute
            Set cmdAtt = Nothing
            connAtt.Close
            Set connAtt = Nothing
        Else
            attId = GenerateNextID("ATT", "AttendanceRecords", "AttendanceID", targetDate)
            InsertAttendanceRow attId, empId, targetDate, reqIn, reqOut, "", "", reqBreak, reqWT, attStatus, requestId & "承認新規反映", approverId
        End If
    End If

    ' Update CorrectionRequests status
    Set connReq = OpenExcelWriteConnection("CorrectionRequests", EXCEL_REQ_REL_PATH)
    Set cmdReq = CreateObject("ADODB.Command")
    Set cmdReq.ActiveConnection = connReq
    cmdReq.CommandType = ADODB_CMD_TEXT
    cmdReq.CommandText = "UPDATE " & EXCEL_REQ_RANGE & _
                         " SET ApprovalStatus = ?, ApproverID = ?, ApprovedAt = ?" & _
                         " WHERE RequestID = ?"
    cmdReq.Parameters.Append cmdReq.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, newStatus)
    cmdReq.Parameters.Append cmdReq.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, approverId)
    cmdReq.Parameters.Append cmdReq.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, nowText)
    cmdReq.Parameters.Append cmdReq.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 20, requestId)
    cmdReq.Execute
    Set cmdReq = Nothing
    connReq.Close
    Set connReq = Nothing

    resultMessage = "申請 " & requestId & " を「" & GetApprovalStatusLabel(newStatus) & "」に更新しました。"
    ApproveOrRejectCorrection = True
    Exit Function

CleanupFail:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    ApproveOrRejectCorrection = False
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not connAtt Is Nothing Then connAtt.Close
    If Not connReq Is Nothing Then connReq.Close
    Set rs = Nothing
    Set db = Nothing
    Set cmdAtt = Nothing
    Set connAtt = Nothing
    Set cmdReq = Nothing
    Set connReq = Nothing
    On Error GoTo 0
    resultMessage = "承認処理エラー: " & Err.Description
    ApproveOrRejectCorrection = False
End Function

' 7. Monthly closing lock / unlock (MonthlyClosings.xlsx)
Public Function ExecuteMonthlyClose( _
    ByVal targetYM As String, _
    ByVal operatorId As String, _
    ByVal closeFlag As String, _
    ByVal noteText As String, _
    ByRef resultMessage As String) As Boolean

    Dim conn As Object
    Dim cmd As Object
    Dim nowText As String
    Dim existsRow As Boolean
    Dim cleanYM As String

    On Error GoTo ErrHandler
    cleanYM = Trim$(targetYM)

    If Not IsManagerOrAdmin(operatorId) Then
        resultMessage = "月次締め操作には Manager または Admin 権限が必要です。"
        ExecuteMonthlyClose = False
        Exit Function
    End If

    If closeFlag = CLOSE_LOCKED Then
        If Not ValidateMonthCloseReady(cleanYM, resultMessage) Then
            ExecuteMonthlyClose = False
            Exit Function
        End If
    ElseIf closeFlag = CLOSE_OPEN Then
        If Not ValidateYMFilter(cleanYM, resultMessage) Then
            ExecuteMonthlyClose = False
            Exit Function
        End If
    Else
        resultMessage = "不正な締めフラグです (" & closeFlag & ")。"
        ExecuteMonthlyClose = False
        Exit Function
    End If

    nowText = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
    existsRow = (DCount("TargetYM", "MonthlyClosings", "TargetYM = '" & SqlEscape(cleanYM) & "'") > 0)

    Set conn = OpenExcelWriteConnection("MonthlyClosings", EXCEL_CLS_REL_PATH)
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = ADODB_CMD_TEXT

    If existsRow Then
        cmd.CommandText = "UPDATE " & EXCEL_CLS_RANGE & _
                          " SET CloseFlag = ?, ClosedBy = ?, ClosedAt = ?, Note = ?" & _
                          " WHERE TargetYM = ?"
        cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, closeFlag)
        cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, IIf(closeFlag = CLOSE_LOCKED, operatorId, ""))
        cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, IIf(closeFlag = CLOSE_LOCKED, nowText, ""))
        cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 255, Trim$(noteText))
        cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 7, cleanYM)
    Else
        cmd.CommandText = "INSERT INTO " & EXCEL_CLS_RANGE & _
                          " (TargetYM, CloseFlag, ClosedBy, ClosedAt, Note)" & _
                          " VALUES (?, ?, ?, ?, ?)"
        cmd.Parameters.Append cmd.CreateParameter("p1", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 7, cleanYM)
        cmd.Parameters.Append cmd.CreateParameter("p2", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 4, closeFlag)
        cmd.Parameters.Append cmd.CreateParameter("p3", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 10, IIf(closeFlag = CLOSE_LOCKED, operatorId, ""))
        cmd.Parameters.Append cmd.CreateParameter("p4", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 19, IIf(closeFlag = CLOSE_LOCKED, nowText, ""))
        cmd.Parameters.Append cmd.CreateParameter("p5", ADODB_TYPE_VARWCHAR, ADODB_PARAM_INPUT, 255, Trim$(noteText))
    End If

    cmd.Execute
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    If closeFlag = CLOSE_LOCKED Then
        resultMessage = cleanYM & " の月次締めを確定しました。"
    Else
        resultMessage = cleanYM & " の月次締めを解除しました。"
    End If
    ExecuteMonthlyClose = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    resultMessage = "月次締め更新エラー: " & Err.Description
    ExecuteMonthlyClose = False
End Function
