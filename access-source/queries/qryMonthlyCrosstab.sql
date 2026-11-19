PARAMETERS [Forms]![frmAttendance]![txtTargetYM] Text ( 7 );
TRANSFORM First([qryAttendanceCalc].[WorkTypeLabel]) AS DayMark
SELECT
    qryAttendanceCalc.EmployeeID,
    qryAttendanceCalc.EmployeeName,
    qryAttendanceCalc.Department,
    Count(qryAttendanceCalc.AttendanceID) AS TotalDays
FROM qryAttendanceCalc
WHERE qryAttendanceCalc.WorkYM = [Forms]![frmAttendance]![txtTargetYM]
GROUP BY
    qryAttendanceCalc.EmployeeID,
    qryAttendanceCalc.EmployeeName,
    qryAttendanceCalc.Department
ORDER BY qryAttendanceCalc.EmployeeID
PIVOT qryAttendanceCalc.DayNum IN (
    '01','02','03','04','05','06','07','08','09','10',
    '11','12','13','14','15','16','17','18','19','20',
    '21','22','23','24','25','26','27','28','29','30','31'
);
