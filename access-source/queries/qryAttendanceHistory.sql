SELECT
    qryAttendanceCalc.AttendanceID,
    qryAttendanceCalc.EmployeeID,
    qryAttendanceCalc.EmployeeName,
    qryAttendanceCalc.Department,
    qryAttendanceCalc.WorkDate,
    qryAttendanceCalc.WorkYM,
    qryAttendanceCalc.ClockIn,
    qryAttendanceCalc.ClockOut,
    qryAttendanceCalc.BreakMinutes,
    qryAttendanceCalc.WorkType,
    qryAttendanceCalc.WorkTypeLabel,
    qryAttendanceCalc.Status,
    qryAttendanceCalc.CloseFlag,
    qryAttendanceCalc.NetWorkMin,
    qryAttendanceCalc.NetWorkHHMM,
    qryAttendanceCalc.OvertimeMin,
    qryAttendanceCalc.LateNightMin,
    qryAttendanceCalc.LateMin,
    qryAttendanceCalc.EarlyLeaveMin,
    qryAttendanceCalc.Remarks
FROM qryAttendanceCalc
WHERE
    (Nz([Forms]![frmAttendance]![txtTargetYM], '') = '' OR qryAttendanceCalc.WorkYM = [Forms]![frmAttendance]![txtTargetYM])
    AND (Nz([Forms]![frmAttendance]![cboFilterEmployee], '') = '' OR qryAttendanceCalc.EmployeeID = [Forms]![frmAttendance]![cboFilterEmployee])
    AND (Nz([Forms]![frmAttendance]![cboFilterStatus], 'All') = 'All' OR qryAttendanceCalc.Status = [Forms]![frmAttendance]![cboFilterStatus])
ORDER BY
    qryAttendanceCalc.WorkDate DESC,
    qryAttendanceCalc.ClockIn DESC;
