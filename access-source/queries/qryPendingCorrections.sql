SELECT
    CorrectionRequests.RequestID,
    CorrectionRequests.AttendanceID,
    CorrectionRequests.EmployeeID,
    Employees.EmployeeName,
    Employees.Department,
    CorrectionRequests.TargetDate,
    CorrectionRequests.ReqClockIn,
    CorrectionRequests.ReqClockOut,
    CorrectionRequests.ReqBreakMin,
    CorrectionRequests.ReqWorkType,
    GetWorkTypeLabel([CorrectionRequests].[ReqWorkType]) AS ReqWorkTypeLabel,
    CorrectionRequests.Reason,
    CorrectionRequests.ApprovalStatus,
    GetApprovalStatusLabel([CorrectionRequests].[ApprovalStatus]) AS ApprovalStatusLabel,
    CorrectionRequests.ApproverID,
    CorrectionRequests.ApprovedAt
FROM CorrectionRequests
INNER JOIN Employees
    ON CorrectionRequests.EmployeeID = Employees.EmployeeID
WHERE CorrectionRequests.ApprovalStatus = '0'
ORDER BY CorrectionRequests.TargetDate DESC, CorrectionRequests.RequestID DESC;
