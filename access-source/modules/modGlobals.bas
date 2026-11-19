Attribute VB_Name = "modGlobals"
Option Compare Database
Option Explicit

' =============================================================================
' Module: modGlobals
' Purpose: Application-wide session state, legacy magic code definitions,
'          TempVars synchronization, OpenArgs parsing, and Excel ADODB helpers.
' History:
'   2011/04/01 山本: 初期版作成 (Access 2007 MDB)
'   2014/09/15 佐藤: TempVars併用に変更 (エラー時のグローバル消失対策)
'   2018/06/01 高橋: Excelリンクテーブル更新制約回避のためADODB接続関数を集約
' =============================================================================

' --- Global runtime variables (legacy; partially mirrored in TempVars) ---
Public gCurrentUserID As String
Public gCurrentRole As String
Public gActiveYM As String
Public gLastErrorMsg As String

' --- Employee status / role constants ---
Public Const EMP_STATUS_ACTIVE As String = "Active"
Public Const EMP_STATUS_INACTIVE As String = "Inactive"
Public Const ROLE_STAFF As String = "Staff"
Public Const ROLE_MANAGER As String = "Manager"
Public Const ROLE_ADMIN As String = "Admin"

' --- Attendance status constants ---
Public Const ATT_STATUS_WORKING As String = "Working"
Public Const ATT_STATUS_ONBREAK As String = "OnBreak"
Public Const ATT_STATUS_COMPLETED As String = "Completed"
Public Const ATT_STATUS_LEAVE As String = "Leave"

' --- Legacy WorkType codes (hardcoded across queries, forms, and payroll CSV) ---
Public Const WT_NORMAL As String = "01"      ' 通常出勤
Public Const WT_LATE_EARLY As String = "02"  ' 遅刻・早退
Public Const WT_HOLIDAY As String = "03"     ' 休日出勤
Public Const WT_PAID_LEAVE As String = "04"  ' 有給休暇
Public Const WT_HALF_LEAVE As String = "05"  ' 半休
Public Const WT_ABSENT As String = "09"      ' 欠勤

' --- Correction request ApprovalStatus codes ---
Public Const APPR_PENDING As String = "0"    ' 申請中
Public Const APPR_APPROVED As String = "1"   ' 承認済
Public Const APPR_REJECTED As String = "2"   ' 差戻

' --- Monthly closing CloseFlag codes ---
Public Const CLOSE_OPEN As String = "0"      ' 未締
Public Const CLOSE_LOCKED As String = "1"    ' 締済

' --- ADODB late-binding constants ---
Public Const ADODB_CMD_TEXT As Long = 1
Public Const ADODB_PARAM_INPUT As Long = 1
Public Const ADODB_TYPE_VARWCHAR As Long = 202

' Called by AutoExec macro via RunCode on database startup.
Public Function InitAppContext() As Boolean
    Dim defaultEmp As Variant

    gActiveYM = Format$(Date, "yyyy-mm")
    defaultEmp = DLookup("EmployeeID", "Employees", "Status = 'Active' AND Role = 'Admin'")
    If IsNull(defaultEmp) Then
        defaultEmp = DMin("EmployeeID", "Employees", "Status = 'Active'")
    End If

    If Not IsNull(defaultEmp) Then
        SetSessionUser CStr(defaultEmp)
    Else
        gCurrentUserID = ""
        gCurrentRole = ROLE_STAFF
        TempVars.Add "ActiveEmployeeID", ""
        TempVars.Add "ActiveRole", ROLE_STAFF
    End If

    TempVars.Add "TargetYM", gActiveYM
    TempVars.Add "AppTitle", "社内勤怠管理システム (Access Legacy Ver 3.4)"
    InitAppContext = True
End Function

' Synchronize both global variables and TempVars when switching operator context.
Public Sub SetSessionUser(ByVal empId As String)
    Dim roleVal As Variant
    Dim cleanId As String

    cleanId = Trim$(empId)
    gCurrentUserID = cleanId
    If Len(cleanId) > 0 Then
        roleVal = DLookup("Role", "Employees", "EmployeeID = '" & SqlEscape(cleanId) & "'")
        gCurrentRole = Nz(roleVal, ROLE_STAFF)
    Else
        gCurrentRole = ROLE_STAFF
    End If

    TempVars.Add "ActiveEmployeeID", gCurrentUserID
    TempVars.Add "ActiveRole", gCurrentRole
End Sub

Public Function GetActiveUserID() As String
    If Len(Trim$(gCurrentUserID)) > 0 Then
        GetActiveUserID = gCurrentUserID
    Else
        GetActiveUserID = Nz(TempVars!ActiveEmployeeID, "")
    End If
End Function

Public Function GetActiveUserRole() As String
    If Len(Trim$(gCurrentRole)) > 0 Then
        GetActiveUserRole = gCurrentRole
    Else
        GetActiveUserRole = Nz(TempVars!ActiveRole, ROLE_STAFF)
    End If
End Function

Public Function IsManagerOrAdmin(ByVal empId As String) As Boolean
    Dim roleVal As Variant

    If Len(Trim$(empId)) = 0 Then
        IsManagerOrAdmin = False
        Exit Function
    End If

    roleVal = DLookup("Role", "Employees", "EmployeeID = '" & SqlEscape(Trim$(empId)) & "' AND Status = 'Active'")
    IsManagerOrAdmin = (Nz(roleVal, "") = ROLE_MANAGER Or Nz(roleVal, "") = ROLE_ADMIN)
End Function

' Decode legacy 2-digit WorkType code. Also invoked from queries and calculated controls.
Public Function GetWorkTypeLabel(ByVal codeVal As Variant) As String
    Select Case Trim$(Nz(codeVal, ""))
        Case WT_NORMAL
            GetWorkTypeLabel = "通常"
        Case WT_LATE_EARLY
            GetWorkTypeLabel = "遅早"
        Case WT_HOLIDAY
            GetWorkTypeLabel = "休出"
        Case WT_PAID_LEAVE
            GetWorkTypeLabel = "有給"
        Case WT_HALF_LEAVE
            GetWorkTypeLabel = "半休"
        Case WT_ABSENT
            GetWorkTypeLabel = "欠勤"
        Case Else
            GetWorkTypeLabel = "不明(" & Nz(codeVal, "") & ")"
    End Select
End Function

' Decode ApprovalStatus numeric text code. Also invoked from queries/forms.
Public Function GetApprovalStatusLabel(ByVal codeVal As Variant) As String
    Select Case Trim$(Nz(codeVal, ""))
        Case APPR_PENDING
            GetApprovalStatusLabel = "申請中"
        Case APPR_APPROVED
            GetApprovalStatusLabel = "承認済"
        Case APPR_REJECTED
            GetApprovalStatusLabel = "差戻"
        Case Else
            GetApprovalStatusLabel = "-"
    End Select
End Function

' Parse pipe-delimited OpenArgs strings, e.g. "EMP001|2026-11|EDIT".
Public Function ParseOpenArgToken(ByVal openArgsVal As Variant, ByVal tokenIndex As Long) As String
    Dim rawText As String
    Dim parts() As String

    rawText = Trim$(Nz(openArgsVal, ""))
    If Len(rawText) = 0 Or tokenIndex < 0 Then
        ParseOpenArgToken = ""
        Exit Function
    End If

    parts = Split(rawText, "|")
    If tokenIndex <= UBound(parts) Then
        ParseOpenArgToken = Trim$(parts(tokenIndex))
    Else
        ParseOpenArgToken = ""
    End If
End Function

' Escape single quotes for Access SQL / Domain Aggregate criteria strings.
Public Function SqlEscape(ByVal rawValue As String) As String
    SqlEscape = Replace(rawValue, "'", "''")
End Function

' Resolve physical Excel workbook path from linked TableDef.Connect or relative fallback.
Public Function GetWorkbookPathByTable(ByVal tableName As String, ByVal fallbackRelPath As String) As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim connectStr As String
    Dim dbPos As Long
    Dim semiPos As Long

    On Error GoTo FallbackPath
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    connectStr = tdf.Connect

    dbPos = InStr(1, connectStr, "DATABASE=", vbTextCompare)
    If dbPos > 0 Then
        dbPos = dbPos + Len("DATABASE=")
        semiPos = InStr(dbPos, connectStr, ";")
        If semiPos > 0 Then
            GetWorkbookPathByTable = Mid$(connectStr, dbPos, semiPos - dbPos)
        Else
            GetWorkbookPathByTable = Mid$(connectStr, dbPos)
        End If
        If Len(Trim$(GetWorkbookPathByTable)) > 0 Then
            Set tdf = Nothing
            Set db = Nothing
            Exit Function
        End If
    End If

FallbackPath:
    Set tdf = Nothing
    Set db = Nothing
    GetWorkbookPathByTable = CurrentProject.Path & "\" & fallbackRelPath
End Function

' Open writable ADODB connection to Excel workbook (workaround for KB 904953 DAO read-only limit).
Public Function OpenExcelWriteConnection(ByVal tableName As String, ByVal fallbackRelPath As String) As Object
    Dim conn As Object
    Dim excelPath As String
    Dim connStr As String

    excelPath = GetWorkbookPathByTable(tableName, fallbackRelPath)
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & excelPath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=YES;READONLY=FALSE;"";"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open connStr
    Set OpenExcelWriteConnection = conn
End Function
