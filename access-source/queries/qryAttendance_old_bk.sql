-- [DEPRECATED / UNUSED OBJECT - KEPT FROM 2014 REFACTORING]
-- Replaced by qryMonthlySummary.sql (which applies 15-min rounding and auto-break via modCalc).
SELECT
    AttendanceRecords.EmployeeID,
    Left(AttendanceRecords.WorkDate, 7) AS YM,
    Count(AttendanceRecords.AttendanceID) AS Cnt
FROM AttendanceRecords
WHERE AttendanceRecords.Status = 'Completed'
GROUP BY
    AttendanceRecords.EmployeeID,
    Left(AttendanceRecords.WorkDate, 7);
