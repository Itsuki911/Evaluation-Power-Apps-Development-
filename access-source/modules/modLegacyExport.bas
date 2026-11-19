Attribute VB_Name = "modLegacyExport"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modLegacyExport
' Purpose: Legacy payroll CSV exporter (fixed-column layout used by old HR
'          payroll package) plus deprecated rounding helper left over from 2012.
' History:
'   2011/06/10 山本: 旧給与ソフト(PCA系)連携CSV出力を作成
'   2012/04/10 佐藤: 10分丸め廃止に伴い LegacyRound10Min_OLD は未使用化
' =============================================================================

' Export monthly attendance summary to legacy fixed-order CSV.
' Called from frmMonthlyClose (btnExportPayrollCsv).
Public Function ExportMonthlyPayrollCsv(ByVal targetYM As String, ByVal outputFilePath As String, ByRef resultMessage As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fileNum As Integer
    Dim sql As String
    Dim lineText As String
    Dim rowCount As Long

    On Error GoTo ErrHandler
    If Not ValidateYMFilter(targetYM, resultMessage) Then
        ExportMonthlyPayrollCsv = False
        Exit Function
    End If

    Set db = CurrentDb
    sql = "SELECT EmployeeID, EmployeeName, Department, WorkYM, WorkDays, " & _
          "TotalWorkMin, TotalOvertimeMin, TotalLateNightMin, LateOrEarlyCount, PaidLeaveDays " & _
          "FROM qryMonthlySummary " & _
          "WHERE WorkYM = '" & SqlEscape(Trim$(targetYM)) & "' " & _
          "ORDER BY EmployeeID ASC;"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    fileNum = FreeFile
    Open outputFilePath For Output As #fileNum

    Print #fileNum, "社員番号,氏名,部署,対象年月,出勤日数,実働分,残業分,深夜分,遅刻早退回数,有給日数"

    rowCount = 0
    Do While Not rs.EOF
        lineText = CStr(Nz(rs!EmployeeID, "")) & "," & _
                   CStr(Nz(rs!EmployeeName, "")) & "," & _
                   CStr(Nz(rs!Department, "")) & "," & _
                   CStr(Nz(rs!WorkYM, "")) & "," & _
                   CStr(Nz(rs!WorkDays, 0)) & "," & _
                   CStr(Nz(rs!TotalWorkMin, 0)) & "," & _
                   CStr(Nz(rs!TotalOvertimeMin, 0)) & "," & _
                   CStr(Nz(rs!TotalLateNightMin, 0)) & "," & _
                   CStr(Nz(rs!LateOrEarlyCount, 0)) & "," & _
                   CStr(Nz(rs!PaidLeaveDays, 0))
        Print #fileNum, lineText
        rowCount = rowCount + 1
        rs.MoveNext
    Loop

    Close #fileNum
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    resultMessage = "給与連携CSVを出力しました (" & rowCount & "件): " & outputFilePath
    ExportMonthlyPayrollCsv = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    resultMessage = "CSV出力エラー: " & Err.Description
    ExportMonthlyPayrollCsv = False
End Function

' [DEPRECATED - UNUSED SINCE 2012/04]
' Retained for historical reference only. Active logic uses modCalc.RoundClockInMinutes (15m).
Public Function LegacyRound10Min_OLD(ByVal rawMinutes As Long) As Long
    LegacyRound10Min_OLD = (rawMinutes \ 10) * 10
End Function
