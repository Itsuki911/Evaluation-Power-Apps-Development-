SELECT
    qryAttendanceBase.AttendanceID,
    qryAttendanceBase.EmployeeID,
    qryAttendanceBase.EmployeeName,
    qryAttendanceBase.Department,
    qryAttendanceBase.PatternCode,
    qryAttendanceBase.WorkDate,
    qryAttendanceBase.WorkYM,
    qryAttendanceBase.DayNum,
    qryAttendanceBase.ClockIn,
    qryAttendanceBase.ClockOut,
    qryAttendanceBase.BreakStart,
    qryAttendanceBase.BreakEnd,
    qryAttendanceBase.BreakMinutes,
    qryAttendanceBase.WorkType,
    GetWorkTypeLabel([qryAttendanceBase].[WorkType]) AS WorkTypeLabel,
    qryAttendanceBase.Status,
    qryAttendanceBase.CloseFlag,
    qryAttendanceBase.Remarks,
    qryAttendanceBase.UpdatedBy,
    CalcWorkMinutes([WorkDate], [ClockIn], [ClockOut], [BreakStart], [BreakEnd], [BreakMinutes], [RoundUnitMin]) AS NetWorkMin,
    FormatMinutesHHMM(CalcWorkMinutes([WorkDate], [ClockIn], [ClockOut], [BreakStart], [BreakEnd], [BreakMinutes], [RoundUnitMin])) AS NetWorkHHMM,
    CalcOvertimeMinutes(
        CalcWorkMinutes([WorkDate], [ClockIn], [ClockOut], [BreakStart], [BreakEnd], [BreakMinutes], [RoundUnitMin]),
        [StdWorkMin]
    ) AS OvertimeMin,
    CalcLateNightMinutes([WorkDate], [ClockIn], [ClockOut], [RoundUnitMin]) AS LateNightMin,
    JudgeLateMinutes([WorkDate], [ClockIn], [SchedStart]) AS LateMin,
    JudgeEarlyLeaveMinutes([WorkDate], [ClockOut], [SchedEnd]) AS EarlyLeaveMin
FROM qryAttendanceBase
ORDER BY
    qryAttendanceBase.WorkDate DESC,
    qryAttendanceBase.ClockIn DESC,
    qryAttendanceBase.AttendanceID DESC;
