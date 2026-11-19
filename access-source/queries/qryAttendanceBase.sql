SELECT
    AttendanceRecords.AttendanceID,
    AttendanceRecords.EmployeeID,
    Employees.EmployeeName,
    Employees.Department,
    Employees.PatternCode,
    WorkPatterns.PatternName,
    WorkPatterns.StartTime AS SchedStart,
    WorkPatterns.EndTime AS SchedEnd,
    Val(Nz(WorkPatterns.StdWorkMin, '480')) AS StdWorkMin,
    Val(Nz(WorkPatterns.RoundUnitMin, '15')) AS RoundUnitMin,
    AttendanceRecords.WorkDate,
    Left(AttendanceRecords.WorkDate, 7) AS WorkYM,
    Right(AttendanceRecords.WorkDate, 2) AS DayNum,
    AttendanceRecords.ClockIn,
    AttendanceRecords.ClockOut,
    AttendanceRecords.BreakStart,
    AttendanceRecords.BreakEnd,
    AttendanceRecords.BreakMinutes,
    AttendanceRecords.WorkType,
    AttendanceRecords.Status,
    AttendanceRecords.Remarks,
    AttendanceRecords.UpdatedBy,
    Nz(MonthlyClosings.CloseFlag, '0') AS CloseFlag
FROM ((AttendanceRecords
INNER JOIN Employees
    ON AttendanceRecords.EmployeeID = Employees.EmployeeID)
LEFT JOIN WorkPatterns
    ON Employees.PatternCode = WorkPatterns.PatternCode)
LEFT JOIN MonthlyClosings
    ON Left(AttendanceRecords.WorkDate, 7) = MonthlyClosings.TargetYM;
