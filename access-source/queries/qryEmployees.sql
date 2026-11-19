SELECT
    Employees.EmployeeID,
    Employees.EmployeeName,
    Employees.Department,
    Employees.Status,
    Employees.Role,
    Employees.PatternCode,
    WorkPatterns.PatternName,
    WorkPatterns.StartTime,
    WorkPatterns.EndTime,
    Employees.ManagerID,
    Employees.HireDate
FROM Employees
LEFT JOIN WorkPatterns
    ON Employees.PatternCode = WorkPatterns.PatternCode
ORDER BY Employees.EmployeeID ASC;
