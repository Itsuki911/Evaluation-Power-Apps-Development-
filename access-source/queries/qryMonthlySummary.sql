SELECT
    qryAttendanceCalc.EmployeeID,
    qryAttendanceCalc.EmployeeName,
    qryAttendanceCalc.Department,
    qryAttendanceCalc.WorkYM,
    Sum(IIf([qryAttendanceCalc].[WorkType] IN ('01', '02', '03', '05') AND [qryAttendanceCalc].[Status] = 'Completed', 1, 0)) AS WorkDays,
    Sum([qryAttendanceCalc].[NetWorkMin]) AS TotalWorkMin,
    Sum([qryAttendanceCalc].[OvertimeMin]) AS TotalOvertimeMin,
    Sum([qryAttendanceCalc].[LateNightMin]) AS TotalLateNightMin,
    Sum(IIf([qryAttendanceCalc].[LateMin] > 0 OR [qryAttendanceCalc].[EarlyLeaveMin] > 0 OR [qryAttendanceCalc].[WorkType] = '02', 1, 0)) AS LateOrEarlyCount,
    Sum(IIf([qryAttendanceCalc].[WorkType] = '04', 1, 0)) AS PaidLeaveDays,
    Sum(IIf([qryAttendanceCalc].[WorkType] = '03', 1, 0)) AS HolidayWorkDays
FROM qryAttendanceCalc
GROUP BY
    qryAttendanceCalc.EmployeeID,
    qryAttendanceCalc.EmployeeName,
    qryAttendanceCalc.Department,
    qryAttendanceCalc.WorkYM
ORDER BY
    qryAttendanceCalc.WorkYM DESC,
    qryAttendanceCalc.EmployeeID ASC;
