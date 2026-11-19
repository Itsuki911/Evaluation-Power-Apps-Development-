SELECT
    AttendanceRecords.AttendanceID,
    AttendanceRecords.EmployeeID,
    Employees.EmployeeName,
    Employees.Department,
    AttendanceRecords.WorkDate,
    AttendanceRecords.ClockIn,
    AttendanceRecords.Status
FROM AttendanceRecords
INNER JOIN Employees
    ON AttendanceRecords.EmployeeID = Employees.EmployeeID
WHERE
    AttendanceRecords.Status = 'Working'
    AND (AttendanceRecords.ClockOut IS NULL OR AttendanceRecords.ClockOut = '')
ORDER BY
    AttendanceRecords.WorkDate DESC,
    AttendanceRecords.ClockIn DESC;
