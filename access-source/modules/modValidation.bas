Attribute VB_Name = "modValidation"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modValidation
' Purpose: Input validation, attendance state transition guards, monthly
'          closing lock checks, and correction request validation.
' History:
'   2011/04/01 山本: 初期作成
'   2016/02/12 佐藤: 休憩開始・終了(OnBreak)の状態遷移チェックを追加
'   2018/11/20 高橋: 月次締め済(CloseFlag='1')の打刻・直接編集ブロックを追加
' =============================================================================

Public Function ValidateEmployeeSelected(ByVal employeeId As String, ByRef errMessage As String) As Boolean
    If Len(Trim$(employeeId)) = 0 Then
        errMessage = "従業員を選択してください。"
        ValidateEmployeeSelected = False
        Exit Function
    End If

    errMessage = ""
    ValidateEmployeeSelected = True
End Function

Public Function ValidateEmployeeActive(ByVal employeeId As String, ByRef errMessage As String) As Boolean
    Dim empStatus As Variant
    Dim criteria As String

    criteria = "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "'"
    empStatus = DLookup("Status", "Employees", criteria)

    If IsNull(empStatus) Then
        errMessage = "指定された従業員が存在しません (" & employeeId & ")。"
        ValidateEmployeeActive = False
        Exit Function
    End If

    If CStr(empStatus) <> EMP_STATUS_ACTIVE Then
        errMessage = "在籍中（Active）ではない従業員は操作できません (" & employeeId & ")。"
        ValidateEmployeeActive = False
        Exit Function
    End If

    errMessage = ""
    ValidateEmployeeActive = True
End Function

' Return True if targetYM ("YYYY-MM") is closed (CloseFlag = '1') in MonthlyClosings.
Public Function IsMonthClosed(ByVal targetYM As String) As Boolean
    Dim flagVal As Variant
    Dim cleanYM As String

    cleanYM = Left$(Trim$(targetYM), 7)
    If Len(cleanYM) <> 7 Then
        IsMonthClosed = False
        Exit Function
    End If

    flagVal = DLookup("CloseFlag", "MonthlyClosings", "TargetYM = '" & SqlEscape(cleanYM) & "'")
    IsMonthClosed = (Trim$(Nz(flagVal, CLOSE_OPEN)) = CLOSE_LOCKED)
End Function

' Return True if employee has an open record (Working or OnBreak with blank ClockOut).
Public Function HasOpenAttendance(ByVal employeeId As String, Optional ByVal workDate As String = "") As Boolean
    Dim criteria As String

    criteria = "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "' " & _
               "AND Status IN ('" & ATT_STATUS_WORKING & "', '" & ATT_STATUS_ONBREAK & "') " & _
               "AND (ClockOut IS NULL OR ClockOut = '')"

    If Len(Trim$(workDate)) > 0 Then
        criteria = criteria & " AND WorkDate = '" & SqlEscape(Trim$(workDate)) & "'"
    End If

    HasOpenAttendance = (DCount("AttendanceID", "AttendanceRecords", criteria) > 0)
End Function

' Return True if employee is currently in OnBreak state.
Public Function HasOnBreakAttendance(ByVal employeeId As String) As Boolean
    Dim criteria As String

    criteria = "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "' " & _
               "AND Status = '" & ATT_STATUS_ONBREAK & "' " & _
               "AND (ClockOut IS NULL OR ClockOut = '')"

    HasOnBreakAttendance = (DCount("AttendanceID", "AttendanceRecords", criteria) > 0)
End Function

' Validate preconditions before Clock-In.
Public Function ValidateClockIn(ByVal employeeId As String, ByVal workDate As String, ByRef errMessage As String) As Boolean
    Dim normalizedId As String
    Dim targetYM As String

    normalizedId = Trim$(employeeId)
    targetYM = Left$(Trim$(workDate), 7)

    If Not ValidateEmployeeSelected(normalizedId, errMessage) Then
        ValidateClockIn = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(normalizedId, errMessage) Then
        ValidateClockIn = False
        Exit Function
    End If

    If IsMonthClosed(targetYM) Then
        errMessage = "対象月 (" & targetYM & ") は既に月次締済のため、直接打刻できません。修正申請を利用してください。"
        ValidateClockIn = False
        Exit Function
    End If

    If HasOpenAttendance(normalizedId, workDate) Then
        errMessage = "同日 (" & workDate & ") に未退勤の出勤記録が既に存在します。重複して出勤打刻はできません。"
        ValidateClockIn = False
        Exit Function
    End If

    If HasOpenAttendance(normalizedId) Then
        errMessage = "未退勤または休憩中の記録が残っています。先に退勤打刻を完了してください。"
        ValidateClockIn = False
        Exit Function
    End If

    errMessage = ""
    ValidateClockIn = True
End Function

' Validate preconditions before Break Start (must be Working, not OnBreak).
Public Function ValidateBreakStart(ByVal employeeId As String, ByVal workDate As String, ByRef errMessage As String) As Boolean
    Dim normalizedId As String

    normalizedId = Trim$(employeeId)
    If Not ValidateEmployeeSelected(normalizedId, errMessage) Then
        ValidateBreakStart = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(normalizedId, errMessage) Then
        ValidateBreakStart = False
        Exit Function
    End If

    If IsMonthClosed(Left$(Trim$(workDate), 7)) Then
        errMessage = "対象月は月次締済のため休憩打刻できません。"
        ValidateBreakStart = False
        Exit Function
    End If

    If HasOnBreakAttendance(normalizedId) Then
        errMessage = "既に休憩中です。先に休憩終了を打刻してください。"
        ValidateBreakStart = False
        Exit Function
    End If

    If Not HasOpenAttendance(normalizedId) Then
        errMessage = "出勤中の記録が見つかりません。先に出勤打刻を行ってください。"
        ValidateBreakStart = False
        Exit Function
    End If

    errMessage = ""
    ValidateBreakStart = True
End Function

' Validate preconditions before Break End (must be OnBreak).
Public Function ValidateBreakEnd(ByVal employeeId As String, ByVal workDate As String, ByRef errMessage As String) As Boolean
    Dim normalizedId As String

    normalizedId = Trim$(employeeId)
    If Not ValidateEmployeeSelected(normalizedId, errMessage) Then
        ValidateBreakEnd = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(normalizedId, errMessage) Then
        ValidateBreakEnd = False
        Exit Function
    End If

    If IsMonthClosed(Left$(Trim$(workDate), 7)) Then
        errMessage = "対象月は月次締済のため休憩終了打刻できません。"
        ValidateBreakEnd = False
        Exit Function
    End If

    If Not HasOnBreakAttendance(normalizedId) Then
        errMessage = "休憩中の記録が見つかりません。先に休憩開始を打刻してください。"
        ValidateBreakEnd = False
        Exit Function
    End If

    errMessage = ""
    ValidateBreakEnd = True
End Function

' Validate preconditions before Clock-Out.
Public Function ValidateClockOut(ByVal employeeId As String, ByRef errMessage As String) As Boolean
    Dim normalizedId As String

    normalizedId = Trim$(employeeId)
    If Not ValidateEmployeeSelected(normalizedId, errMessage) Then
        ValidateClockOut = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(normalizedId, errMessage) Then
        ValidateClockOut = False
        Exit Function
    End If

    If HasOnBreakAttendance(normalizedId) Then
        errMessage = "現在休憩中です。退勤打刻の前に「休憩終了」を打刻してください。"
        ValidateClockOut = False
        Exit Function
    End If

    If Not HasOpenAttendance(normalizedId) Then
        errMessage = "未退勤の出勤記録が見つかりません。先に出勤打刻を行ってください。"
        ValidateClockOut = False
        Exit Function
    End If

    errMessage = ""
    ValidateClockOut = True
End Function

' Validate correction request submission input.
Public Function ValidateCorrectionInput( _
    ByVal employeeId As String, _
    ByVal targetDate As String, _
    ByVal reqClockIn As String, _
    ByVal reqClockOut As String, _
    ByVal reqBreakMin As String, _
    ByVal reqWorkType As String, _
    ByVal reasonText As String, _
    ByRef errMessage As String) As Boolean

    If Not ValidateEmployeeSelected(employeeId, errMessage) Then
        ValidateCorrectionInput = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(employeeId, errMessage) Then
        ValidateCorrectionInput = False
        Exit Function
    End If

    If Len(Trim$(targetDate)) = 0 Or Not ValidateDateFilter(targetDate, errMessage) Then
        If Len(errMessage) = 0 Then errMessage = "対象日を入力してください (YYYY-MM-DD)。"
        ValidateCorrectionInput = False
        Exit Function
    End If

    If Len(Trim$(reasonText)) = 0 Then
        errMessage = "申請理由を入力してください。"
        ValidateCorrectionInput = False
        Exit Function
    End If

    Select Case Trim$(reqWorkType)
        Case WT_NORMAL, WT_LATE_EARLY, WT_HOLIDAY, WT_PAID_LEAVE, WT_HALF_LEAVE, WT_ABSENT
            ' valid legacy code
        Case Else
            errMessage = "勤務区分コードが不正です (" & reqWorkType & ")。"
            ValidateCorrectionInput = False
            Exit Function
    End Select

    If reqWorkType <> WT_PAID_LEAVE And reqWorkType <> WT_ABSENT Then
        If Len(Trim$(reqClockIn)) < 16 Or Not IsDate(reqClockIn) Then
            errMessage = "修正後出勤時刻は YYYY-MM-DD HH:MM:SS 形式で入力してください。"
            ValidateCorrectionInput = False
            Exit Function
        End If
        If Len(Trim$(reqClockOut)) < 16 Or Not IsDate(reqClockOut) Then
            errMessage = "修正後退勤時刻は YYYY-MM-DD HH:MM:SS 形式で入力してください。"
            ValidateCorrectionInput = False
            Exit Function
        End If
        If CDate(reqClockOut) <= CDate(reqClockIn) Then
            errMessage = "修正後退勤時刻は出勤時刻より後に設定してください。"
            ValidateCorrectionInput = False
            Exit Function
        End If
    End If

    If Len(Trim$(reqBreakMin)) > 0 Then
        If Not IsNumeric(reqBreakMin) Or CLng(reqBreakMin) < 0 Then
            errMessage = "休憩分数は0以上の数値で入力してください。"
            ValidateCorrectionInput = False
            Exit Function
        End If
    End If

    errMessage = ""
    ValidateCorrectionInput = True
End Function

' Verify whether targetYM ("YYYY-MM") can be closed without blocking anomalies.
Public Function ValidateMonthCloseReady(ByVal targetYM As String, ByRef errMessage As String) As Boolean
    Dim openCount As Long
    Dim pendingReqCount As Long
    Dim safeYM As String

    If Not ValidateYMFilter(targetYM, errMessage) Then
        ValidateMonthCloseReady = False
        Exit Function
    End If

    safeYM = SqlEscape(Trim$(targetYM))

    openCount = DCount("AttendanceID", "AttendanceRecords", _
        "Left(WorkDate, 7) = '" & safeYM & "' AND Status IN ('" & ATT_STATUS_WORKING & "', '" & ATT_STATUS_ONBREAK & "')")
    If openCount > 0 Then
        errMessage = targetYM & " に未退勤・休憩中のレコードが " & openCount & " 件残っているため締め処理できません。"
        ValidateMonthCloseReady = False
        Exit Function
    End If

    pendingReqCount = DCount("RequestID", "CorrectionRequests", _
        "Left(TargetDate, 7) = '" & safeYM & "' AND ApprovalStatus = '" & APPR_PENDING & "'")
    If pendingReqCount > 0 Then
        errMessage = targetYM & " に未承認の修正申請が " & pendingReqCount & " 件残っているため締め処理できません。"
        ValidateMonthCloseReady = False
        Exit Function
    End If

    errMessage = ""
    ValidateMonthCloseReady = True
End Function

' Validate optional WorkDate input ("YYYY-MM-DD").
Public Function ValidateDateFilter(ByVal dateText As String, ByRef errMessage As String) As Boolean
    Dim trimmed As String
    Dim parsedDate As Date

    trimmed = Trim$(dateText)
    If Len(trimmed) = 0 Then
        errMessage = ""
        ValidateDateFilter = True
        Exit Function
    End If

    If Len(trimmed) <> 10 Or Mid$(trimmed, 5, 1) <> "-" Or Mid$(trimmed, 8, 1) <> "-" Then
        errMessage = "日付は YYYY-MM-DD 形式で入力してください（例: 2026-11-18）。"
        ValidateDateFilter = False
        Exit Function
    End If

    If Not IsNumeric(Left$(trimmed, 4)) Or Not IsNumeric(Mid$(trimmed, 6, 2)) Or Not IsNumeric(Right$(trimmed, 2)) Then
        errMessage = "日付は YYYY-MM-DD 形式で入力してください（例: 2026-11-18）。"
        ValidateDateFilter = False
        Exit Function
    End If

    If Not IsDate(trimmed) Then
        errMessage = "存在しない日付が指定されました (" & trimmed & ")。"
        ValidateDateFilter = False
        Exit Function
    End If

    parsedDate = CDate(trimmed)
    If Format$(parsedDate, "yyyy-mm-dd") <> trimmed Then
        errMessage = "存在しない日付が指定されました (" & trimmed & ")。"
        ValidateDateFilter = False
        Exit Function
    End If

    errMessage = ""
    ValidateDateFilter = True
End Function

' Validate TargetYM ("YYYY-MM").
Public Function ValidateYMFilter(ByVal ymText As String, ByRef errMessage As String) As Boolean
    Dim trimmed As String

    trimmed = Trim$(ymText)
    If Len(trimmed) <> 7 Or Mid$(trimmed, 5, 1) <> "-" Then
        errMessage = "対象年月は YYYY-MM 形式で入力してください（例: 2026-11）。"
        ValidateYMFilter = False
        Exit Function
    End If

    If Not IsNumeric(Left$(trimmed, 4)) Or Not IsNumeric(Right$(trimmed, 2)) Then
        errMessage = "対象年月は YYYY-MM 形式で入力してください（例: 2026-11）。"
        ValidateYMFilter = False
        Exit Function
    End If

    If CLng(Right$(trimmed, 2)) < 1 Or CLng(Right$(trimmed, 2)) > 12 Then
        errMessage = "対象年月の月部分が不正です (" & trimmed & ")。"
        ValidateYMFilter = False
        Exit Function
    End If

    errMessage = ""
    ValidateYMFilter = True
End Function
