Attribute VB_Name = "modCalc"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modCalc
' Purpose: Labor time calculation functions invoked BOTH by VBA routines AND
'          directly inside Access SQL queries (qryAttendanceCalc) and form
'          calculated controls.
'
' Business rules embedded here:
'   - 15-minute rounding (2012/04/10 佐藤: 就業規則改定に伴い15分丸め導入)
'       Clock-In  : Round UP (ceil) to nearest RoundUnitMin (default 15m)
'       Clock-Out : Round DOWN (floor) to nearest RoundUnitMin (default 15m)
'   - Statutory automatic break deduction when manual break is not recorded:
'       Gross > 480 min (8h) -> 60 min
'       Gross > 360 min (6h) -> 45 min
'       Otherwise            -> 0 min
'   - Overtime: Actual worked minutes exceeding StdWorkMin (default 480 min)
'   - Late-night (2015/08/03 鈴木): Overlap with 22:00-05:00 (including
'       shifts crossing midnight up to next-day 05:00)
' =============================================================================

Public Const LATE_NIGHT_START_MIN As Long = 1320   ' 22:00 = 22 * 60
Public Const LATE_NIGHT_END_MIN As Long = 1740     ' Next day 05:00 = 29 * 60
Public Const EARLY_MORNING_END_MIN As Long = 300   ' 05:00 = 5 * 60

' Parse "YYYY-MM-DD HH:MM:SS" relative to WorkDate ("YYYY-MM-DD") into minutes from 00:00 of WorkDate.
' Supports overnight shifts (ClockOut on WorkDate + 1 returns >= 1440).
Public Function ParseOffsetMinutes(ByVal workDateText As Variant, ByVal dateTimeText As Variant) As Long
    Dim baseDt As Date
    Dim targetDt As Date
    Dim wd As String
    Dim dtStr As String

    wd = Trim$(Nz(workDateText, ""))
    dtStr = Trim$(Nz(dateTimeText, ""))
    If Len(wd) < 10 Or Len(dtStr) < 16 Then
        ParseOffsetMinutes = -1
        Exit Function
    End If

    If Not IsDate(wd) Or Not IsDate(dtStr) Then
        ParseOffsetMinutes = -1
        Exit Function
    End If

    baseDt = CDate(wd)
    targetDt = CDate(dtStr)
    ParseOffsetMinutes = DateDiff("n", baseDt, targetDt)
End Function

' Parse "HH:MM" schedule string into minutes from midnight.
Public Function ParseHHMMToMinutes(ByVal hhmmText As Variant) As Long
    Dim s As String
    Dim parts() As String

    s = Trim$(Nz(hhmmText, ""))
    If Len(s) < 4 Or InStr(1, s, ":") = 0 Then
        ParseHHMMToMinutes = -1
        Exit Function
    End If

    parts = Split(s, ":")
    If UBound(parts) < 1 Then
        ParseHHMMToMinutes = -1
        Exit Function
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
        ParseHHMMToMinutes = -1
        Exit Function
    End If

    ParseHHMMToMinutes = CLng(parts(0)) * 60 + CLng(parts(1))
End Function

' Round Clock-In UP (ceil) to unitMin boundary (e.g., 09:03 -> 09:15, 09:00 -> 09:00).
Public Function RoundClockInMinutes(ByVal workDateText As Variant, ByVal clockInText As Variant, Optional ByVal unitMin As Long = 15) As Long
    Dim rawMin As Long

    rawMin = ParseOffsetMinutes(workDateText, clockInText)
    If rawMin < 0 Then
        RoundClockInMinutes = -1
        Exit Function
    End If

    If unitMin <= 1 Then
        RoundClockInMinutes = rawMin
    Else
        RoundClockInMinutes = ((rawMin + unitMin - 1) \ unitMin) * unitMin
    End If
End Function

' Round Clock-Out DOWN (floor) to unitMin boundary (e.g., 18:12 -> 18:00, 18:15 -> 18:15).
Public Function RoundClockOutMinutes(ByVal workDateText As Variant, ByVal clockOutText As Variant, Optional ByVal unitMin As Long = 15) As Long
    Dim rawMin As Long

    rawMin = ParseOffsetMinutes(workDateText, clockOutText)
    If rawMin < 0 Then
        RoundClockOutMinutes = -1
        Exit Function
    End If

    If unitMin <= 1 Then
        RoundClockOutMinutes = rawMin
    Else
        RoundClockOutMinutes = (rawMin \ unitMin) * unitMin
    End If
End Function

' Determine effective break minutes:
'   1) Explicit BreakMinutes if stored
'   2) Actual BreakEnd - BreakStart if both recorded
'   3) Statutory automatic deduction (>8h: 60m, >6h: 45m, else 0m)
Public Function CalcEffectiveBreakMinutes( _
    ByVal grossMin As Long, _
    ByVal breakStartText As Variant, _
    ByVal breakEndText As Variant, _
    ByVal explicitBreakMin As Variant) As Long

    Dim expStr As String
    Dim bs As String
    Dim be As String
    Dim diffMin As Long

    If grossMin <= 0 Then
        CalcEffectiveBreakMinutes = 0
        Exit Function
    End If

    expStr = Trim$(Nz(explicitBreakMin, ""))
    If Len(expStr) > 0 And IsNumeric(expStr) Then
        If CLng(expStr) >= 0 Then
            CalcEffectiveBreakMinutes = CLng(expStr)
            Exit Function
        End If
    End If

    bs = Trim$(Nz(breakStartText, ""))
    be = Trim$(Nz(breakEndText, ""))
    If Len(bs) >= 16 And Len(be) >= 16 And IsDate(bs) And IsDate(be) Then
        diffMin = DateDiff("n", CDate(bs), CDate(be))
        If diffMin >= 0 Then
            CalcEffectiveBreakMinutes = diffMin
            Exit Function
        End If
    End If

    ' Statutory auto-deduction fallback
    If grossMin > 480 Then
        CalcEffectiveBreakMinutes = 60
    ElseIf grossMin > 360 Then
        CalcEffectiveBreakMinutes = 45
    Else
        CalcEffectiveBreakMinutes = 0
    End If
End Function

' Compute net worked minutes after 15-min rounding and break deduction.
' Invoked directly by qryAttendanceCalc.sql.
Public Function CalcWorkMinutes( _
    ByVal workDateText As Variant, _
    ByVal clockInText As Variant, _
    ByVal clockOutText As Variant, _
    ByVal breakStartText As Variant, _
    ByVal breakEndText As Variant, _
    ByVal explicitBreakMin As Variant, _
    Optional ByVal unitMin As Long = 15) As Long

    Dim inMin As Long
    Dim outMin As Long
    Dim grossMin As Long
    Dim breakMin As Long
    Dim netMin As Long

    inMin = RoundClockInMinutes(workDateText, clockInText, unitMin)
    outMin = RoundClockOutMinutes(workDateText, clockOutText, unitMin)
    If inMin < 0 Or outMin < 0 Or outMin <= inMin Then
        CalcWorkMinutes = 0
        Exit Function
    End If

    grossMin = outMin - inMin
    breakMin = CalcEffectiveBreakMinutes(grossMin, breakStartText, breakEndText, explicitBreakMin)
    netMin = grossMin - breakMin
    If netMin < 0 Then netMin = 0

    CalcWorkMinutes = netMin
End Function

' Compute daily overtime minutes beyond standard working minutes (default 480m = 8h).
Public Function CalcOvertimeMinutes(ByVal netWorkMin As Variant, Optional ByVal stdWorkMin As Variant = 480) As Long
    Dim wMin As Long
    Dim sMin As Long

    wMin = CLng(Nz(netWorkMin, 0))
    If IsNumeric(Nz(stdWorkMin, 480)) Then
        sMin = CLng(Nz(stdWorkMin, 480))
    Else
        sMin = 480
    End If

    If wMin > sMin Then
        CalcOvertimeMinutes = wMin - sMin
    Else
        CalcOvertimeMinutes = 0
    End If
End Function

' Compute overlap minutes between rounded shift [inMin, outMin] and late-night windows
' (00:00-05:00 and 22:00-next day 05:00).
Public Function CalcLateNightMinutes( _
    ByVal workDateText As Variant, _
    ByVal clockInText As Variant, _
    ByVal clockOutText As Variant, _
    Optional ByVal unitMin As Long = 15) As Long

    Dim inMin As Long
    Dim outMin As Long
    Dim overlapEarly As Long
    Dim overlapNight As Long

    inMin = RoundClockInMinutes(workDateText, clockInText, unitMin)
    outMin = RoundClockOutMinutes(workDateText, clockOutText, unitMin)
    If inMin < 0 Or outMin < 0 Or outMin <= inMin Then
        CalcLateNightMinutes = 0
        Exit Function
    End If

    overlapEarly = CalcIntervalOverlap(inMin, outMin, 0, EARLY_MORNING_END_MIN)
    overlapNight = CalcIntervalOverlap(inMin, outMin, LATE_NIGHT_START_MIN, LATE_NIGHT_END_MIN)
    CalcLateNightMinutes = overlapEarly + overlapNight
End Function

Private Function CalcIntervalOverlap(ByVal aStart As Long, ByVal aEnd As Long, ByVal bStart As Long, ByVal bEnd As Long) As Long
    Dim s As Long
    Dim e As Long

    If aStart > bStart Then s = aStart Else s = bStart
    If aEnd < bEnd Then e = aEnd Else e = bEnd

    If e > s Then
        CalcIntervalOverlap = e - s
    Else
        CalcIntervalOverlap = 0
    End If
End Function

' Return late arrival minutes (actual unrounded ClockIn > pattern StartTime).
Public Function JudgeLateMinutes(ByVal workDateText As Variant, ByVal clockInText As Variant, ByVal schedStartHHMM As Variant) As Long
    Dim actualIn As Long
    Dim schedIn As Long

    actualIn = ParseOffsetMinutes(workDateText, clockInText)
    schedIn = ParseHHMMToMinutes(schedStartHHMM)
    If actualIn < 0 Or schedIn < 0 Then
        JudgeLateMinutes = 0
        Exit Function
    End If

    If actualIn > schedIn Then
        JudgeLateMinutes = actualIn - schedIn
    Else
        JudgeLateMinutes = 0
    End If
End Function

' Return early departure minutes (actual unrounded ClockOut < pattern EndTime).
Public Function JudgeEarlyLeaveMinutes(ByVal workDateText As Variant, ByVal clockOutText As Variant, ByVal schedEndHHMM As Variant) As Long
    Dim actualOut As Long
    Dim schedOut As Long

    actualOut = ParseOffsetMinutes(workDateText, clockOutText)
    schedOut = ParseHHMMToMinutes(schedEndHHMM)
    If actualOut < 0 Or schedOut < 0 Then
        JudgeEarlyLeaveMinutes = 0
        Exit Function
    End If

    If actualOut < schedOut Then
        JudgeEarlyLeaveMinutes = schedOut - actualOut
    Else
        JudgeEarlyLeaveMinutes = 0
    End If
End Function

' Format total minutes as "HH:MM" (e.g., 510 -> "08:30"). Used in forms and reports.
Public Function FormatMinutesHHMM(ByVal totalMinutes As Variant) As String
    Dim m As Long
    Dim hrs As Long
    Dim mins As Long

    m = CLng(Nz(totalMinutes, 0))
    If m < 0 Then m = 0
    hrs = m \ 60
    mins = m Mod 60
    FormatMinutesHHMM = Format$(hrs, "00") & ":" & Format$(mins, "00")
End Function
