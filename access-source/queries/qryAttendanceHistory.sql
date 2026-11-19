SELECT
    AttendanceRecords.AttendanceID,
    AttendanceRecords.EmployeeID,
    Employees.EmployeeName,
    Employees.Department,
    AttendanceRecords.WorkDate,
    AttendanceRecords.ClockIn,
    AttendanceRecords.ClockOut,
    AttendanceRecords.Status
FROM AttendanceRecords
INNER JOIN Employees
    ON AttendanceRecords.EmployeeID = Employees.EmployeeID
ORDER BY
    AttendanceRecords.WorkDate DESC,
    AttendanceRecords.ClockIn DESC,
    AttendanceRecords.AttendanceID DESC;
