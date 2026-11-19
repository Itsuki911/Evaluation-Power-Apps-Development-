Attribute VB_Name = "modValidation"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modValidation
' Purpose: Input validation and attendance state transition rules
' =============================================================================

Public Const EMPLOYEE_STATUS_ACTIVE As String = "Active"
Public Const ATTENDANCE_STATUS_WORKING As String = "Working"

' Escape single quotes for Access SQL / Domain Aggregate criteria strings.
Public Function SqlEscape(ByVal rawValue As String) As String
    SqlEscape = Replace(rawValue, "'", "''")
End Function

' Verify that an EmployeeID has been selected.
Public Function ValidateEmployeeSelected(ByVal employeeId As String, ByRef errMessage As String) As Boolean
    If Len(Trim$(employeeId)) = 0 Then
        errMessage = "従業員を選択してください。"
        ValidateEmployeeSelected = False
        Exit Function
    End If

    errMessage = ""
    ValidateEmployeeSelected = True
End Function

' Verify that the specified employee exists and has Status = "Active".
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

    If CStr(empStatus) <> EMPLOYEE_STATUS_ACTIVE Then
        errMessage = "在籍中（Active）ではない従業員は打刻できません (" & employeeId & ")。"
        ValidateEmployeeActive = False
        Exit Function
    End If

    errMessage = ""
    ValidateEmployeeActive = True
End Function

' Return True if the employee has an unclosed ("Working" and blank ClockOut) record.
' When workDate is provided ("YYYY-MM-DD"), checks specifically for that date.
Public Function HasOpenAttendance(ByVal employeeId As String, Optional ByVal workDate As String = "") As Boolean
    Dim criteria As String

    criteria = "EmployeeID = '" & SqlEscape(Trim$(employeeId)) & "' " & _
               "AND Status = '" & ATTENDANCE_STATUS_WORKING & "' " & _
               "AND (ClockOut IS NULL OR ClockOut = '')"

    If Len(Trim$(workDate)) > 0 Then
        criteria = criteria & " AND WorkDate = '" & SqlEscape(Trim$(workDate)) & "'"
    End If

    HasOpenAttendance = (DCount("AttendanceID", "AttendanceRecords", criteria) > 0)
End Function

' Validate all preconditions before executing Clock-In.
Public Function ValidateClockIn(ByVal employeeId As String, ByVal workDate As String, ByRef errMessage As String) As Boolean
    Dim normalizedId As String
    normalizedId = Trim$(employeeId)

    If Not ValidateEmployeeSelected(normalizedId, errMessage) Then
        ValidateClockIn = False
        Exit Function
    End If

    If Not ValidateEmployeeActive(normalizedId, errMessage) Then
        ValidateClockIn = False
        Exit Function
    End If

    ' Prevent duplicate clock-in on the same date while still clocked in.
    If HasOpenAttendance(normalizedId, workDate) Then
        errMessage = "同日 (" & workDate & ") に未退勤の出勤記録が既に存在します。重複して出勤打刻はできません。"
        ValidateClockIn = False
        Exit Function
    End If

    ' Prevent opening a second concurrent working shift if a prior day remains unclosed.
    If HasOpenAttendance(normalizedId) Then
        errMessage = "未退勤の出勤記録が残っています。先に退勤打刻を完了してください。"
        ValidateClockIn = False
        Exit Function
    End If

    errMessage = ""
    ValidateClockIn = True
End Function

' Validate all preconditions before executing Clock-Out.
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

    ' Require an existing open attendance record before allowing clock-out.
    If Not HasOpenAttendance(normalizedId) Then
        errMessage = "未退勤の出勤記録が見つかりません。先に出勤打刻を行ってください。"
        ValidateClockOut = False
        Exit Function
    End If

    errMessage = ""
    ValidateClockOut = True
End Function

' Validate optional WorkDate filter input ("YYYY-MM-DD"). Empty string is allowed.
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
        errMessage = "勤務日は YYYY-MM-DD 形式で入力してください（例: 2026-11-18）。"
        ValidateDateFilter = False
        Exit Function
    End If

    If Not IsNumeric(Left$(trimmed, 4)) Or Not IsNumeric(Mid$(trimmed, 6, 2)) Or Not IsNumeric(Right$(trimmed, 2)) Then
        errMessage = "勤務日は YYYY-MM-DD 形式で入力してください（例: 2026-11-18）。"
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
