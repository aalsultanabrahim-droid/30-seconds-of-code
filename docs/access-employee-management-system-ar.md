# Professional Employee Management System for Microsoft Access

> الهدف: بناء نظام متكامل لإدارة الموظفين، الحضور، الساعات الإضافية، السلف/السحبيات، الرواتب، مراكز التكلفة، الصلاحيات، التدقيق، والتقارير الاحترافية.
>
> **لغة الكائنات داخل Access:** English (موصى به للاستقرار والتكامل)، مع تسميات واجهات وتقارير عربية/ثنائية اللغة.

---

## 1) Naming Convention (ثابتة واحترافية)

- **Tables:** `tblEmployees`, `tblAttendance`, `tblWithdrawals`, `tblPayroll`, `tblCostCenters`, `tblEmployeeCostCenter`, `tblUsers`, `tblAuditLog`, `tblSettings`, `tblDeductions`, `tblWeeklyEntitlements`
- **Queries:** `qry...`
- **Forms:** `frm...`
- **Reports:** `rpt...`
- **Macros:** `mcr...`
- **Modules:** `mod...`

> القاعدة: اسم إنجليزي تقني + Caption عربي واضح داخل الواجهة.

---

## 2) Table Design (DDL for Access SQL View)

> افتح Access → Create Query → SQL View، ونفذ الأوامر بالتسلسل (عدّل الأنواع إذا لزم حسب إصدار Access).

```sql
CREATE TABLE tblEmployees (
  EmployeeID AUTOINCREMENT CONSTRAINT PK_tblEmployees PRIMARY KEY,
  EmployeeCode TEXT(20),
  FullName TEXT(150) NOT NULL,
  JobTitle TEXT(100),
  HourlyRate CURRENCY DEFAULT 0,
  Phone TEXT(30),
  NationalID TEXT(30),
  HireDate DATETIME,
  Status TEXT(15) DEFAULT 'Active',
  CreatedAt DATETIME DEFAULT Now(),
  UpdatedAt DATETIME
);

CREATE TABLE tblAttendance (
  AttendanceID AUTOINCREMENT CONSTRAINT PK_tblAttendance PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  WorkDate DATETIME NOT NULL,
  StartTime DATETIME,
  EndTime DATETIME,
  BreakMinutes LONG DEFAULT 0,
  TotalHours DOUBLE,
  OvertimeHours DOUBLE DEFAULT 0,
  AttendanceStatus TEXT(20) DEFAULT 'Present',
  Notes LONGTEXT,
  CreatedAt DATETIME DEFAULT Now(),
  UpdatedAt DATETIME,
  CONSTRAINT FK_Attendance_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID)
);

CREATE TABLE tblWithdrawals (
  WithdrawalID AUTOINCREMENT CONSTRAINT PK_tblWithdrawals PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  WithdrawalDate DATETIME NOT NULL,
  Amount CURRENCY NOT NULL,
  WithdrawalType TEXT(20) DEFAULT 'Advance',
  Description LONGTEXT,
  ApprovedBy LONG,
  CreatedAt DATETIME DEFAULT Now(),
  CONSTRAINT FK_Withdrawals_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID)
);

CREATE TABLE tblDeductions (
  DeductionID AUTOINCREMENT CONSTRAINT PK_tblDeductions PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  DeductionDate DATETIME NOT NULL,
  Amount CURRENCY NOT NULL,
  DeductionReason TEXT(100),
  Notes LONGTEXT,
  CONSTRAINT FK_Deductions_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID)
);

CREATE TABLE tblPayroll (
  PayrollID AUTOINCREMENT CONSTRAINT PK_tblPayroll PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  PeriodStart DATETIME NOT NULL,
  PeriodEnd DATETIME NOT NULL,
  TotalWorkHours DOUBLE DEFAULT 0,
  TotalOvertimeHours DOUBLE DEFAULT 0,
  HourlyRate CURRENCY DEFAULT 0,
  BaseSalary CURRENCY DEFAULT 0,
  OvertimeRate CURRENCY DEFAULT 0,
  OvertimePay CURRENCY DEFAULT 0,
  TotalWithdrawals CURRENCY DEFAULT 0,
  TotalDeductions CURRENCY DEFAULT 0,
  Bonuses CURRENCY DEFAULT 0,
  NetSalary CURRENCY DEFAULT 0,
  PayrollStatus TEXT(20) DEFAULT 'Draft',
  ApprovedBy LONG,
  ApprovedAt DATETIME,
  CreatedAt DATETIME DEFAULT Now(),
  CONSTRAINT FK_Payroll_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID)
);

CREATE TABLE tblCostCenters (
  CostCenterID AUTOINCREMENT CONSTRAINT PK_tblCostCenters PRIMARY KEY,
  CostCenterCode TEXT(20),
  CostCenterName TEXT(120) NOT NULL,
  Description LONGTEXT,
  IsActive YESNO DEFAULT TRUE
);

CREATE TABLE tblEmployeeCostCenter (
  EmployeeCostCenterID AUTOINCREMENT CONSTRAINT PK_tblEmployeeCostCenter PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  CostCenterID LONG NOT NULL,
  AllocationPercent DOUBLE DEFAULT 100,
  AllocatedAmount CURRENCY DEFAULT 0,
  EffectiveFrom DATETIME,
  EffectiveTo DATETIME,
  CONSTRAINT FK_ECC_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID),
  CONSTRAINT FK_ECC_CostCenter FOREIGN KEY (CostCenterID)
    REFERENCES tblCostCenters (CostCenterID)
);

CREATE TABLE tblWeeklyEntitlements (
  EntitlementID AUTOINCREMENT CONSTRAINT PK_tblWeeklyEntitlements PRIMARY KEY,
  EmployeeID LONG NOT NULL,
  WeekStart DATETIME NOT NULL,
  WeekEnd DATETIME NOT NULL,
  GrossEntitlement CURRENCY DEFAULT 0,
  WithdrawalsApplied CURRENCY DEFAULT 0,
  DeductionsApplied CURRENCY DEFAULT 0,
  NetEntitlement CURRENCY DEFAULT 0,
  Notes LONGTEXT,
  CONSTRAINT FK_Entitlement_Employee FOREIGN KEY (EmployeeID)
    REFERENCES tblEmployees (EmployeeID)
);

CREATE TABLE tblUsers (
  UserID AUTOINCREMENT CONSTRAINT PK_tblUsers PRIMARY KEY,
  UserName TEXT(60) NOT NULL,
  UserPasswordHash TEXT(255) NOT NULL,
  FullName TEXT(100),
  UserRole TEXT(20) NOT NULL,
  IsActive YESNO DEFAULT TRUE,
  LastLogin DATETIME
);

CREATE TABLE tblAuditLog (
  AuditID AUTOINCREMENT CONSTRAINT PK_tblAuditLog PRIMARY KEY,
  UserID LONG,
  ActionType TEXT(30),
  TableName TEXT(80),
  RecordID LONG,
  FieldName TEXT(80),
  OldValue LONGTEXT,
  NewValue LONGTEXT,
  ActionTime DATETIME DEFAULT Now(),
  Notes LONGTEXT
);

CREATE TABLE tblSettings (
  SettingID AUTOINCREMENT CONSTRAINT PK_tblSettings PRIMARY KEY,
  SettingKey TEXT(100) NOT NULL,
  SettingValue LONGTEXT,
  UpdatedAt DATETIME DEFAULT Now()
);
```

### Important Indexes + Uniqueness

```sql
CREATE UNIQUE INDEX UX_EmployeeCode ON tblEmployees (EmployeeCode);
CREATE UNIQUE INDEX UX_Attendance_Employee_WorkDate ON tblAttendance (EmployeeID, WorkDate);
CREATE INDEX IX_Attendance_WorkDate ON tblAttendance (WorkDate);
CREATE INDEX IX_Withdrawals_Date ON tblWithdrawals (WithdrawalDate);
CREATE INDEX IX_Payroll_Period ON tblPayroll (PeriodStart, PeriodEnd);
```

---

## 3) Relationships (with Accounting-grade Integrity)

من Database Tools → Relationships:

- `tblEmployees (1) -> (∞) tblAttendance`
- `tblEmployees (1) -> (∞) tblWithdrawals`
- `tblEmployees (1) -> (∞) tblDeductions`
- `tblEmployees (1) -> (∞) tblPayroll`
- `tblEmployees (1) -> (∞) tblEmployeeCostCenter`
- `tblCostCenters (1) -> (∞) tblEmployeeCostCenter`
- `tblEmployees (1) -> (∞) tblWeeklyEntitlements`

إعدادات مطلوبة:

- ✅ Enable Referential Integrity
- ✅ Cascade Update Related Fields
- ⚠️ Cascade Delete: يفضّل **عدم التفعيل** في الرواتب/الحضور للمحافظة على السجل المالي.

---

## 4) Core Business Queries

### 4.1 Total worked hours by period

```sql
PARAMETERS pEmployeeID Long, pDateFrom DateTime, pDateTo DateTime;
SELECT
  a.EmployeeID,
  Sum(Nz(a.TotalHours,0)) AS SumWorkHours,
  Sum(Nz(a.OvertimeHours,0)) AS SumOvertimeHours
FROM tblAttendance AS a
WHERE a.EmployeeID = [pEmployeeID]
  AND a.WorkDate BETWEEN [pDateFrom] AND [pDateTo]
GROUP BY a.EmployeeID;
```

احفظه باسم: `qryEmployeeHoursByPeriod`

### 4.2 Total withdrawals by period

```sql
PARAMETERS pEmployeeID Long, pDateFrom DateTime, pDateTo DateTime;
SELECT
  w.EmployeeID,
  Sum(Nz(w.Amount,0)) AS SumWithdrawals
FROM tblWithdrawals AS w
WHERE w.EmployeeID = [pEmployeeID]
  AND w.WithdrawalDate BETWEEN [pDateFrom] AND [pDateTo]
GROUP BY w.EmployeeID;
```

احفظه باسم: `qryEmployeeWithdrawalsByPeriod`

### 4.3 Total deductions by period

```sql
PARAMETERS pEmployeeID Long, pDateFrom DateTime, pDateTo DateTime;
SELECT
  d.EmployeeID,
  Sum(Nz(d.Amount,0)) AS SumDeductions
FROM tblDeductions AS d
WHERE d.EmployeeID = [pEmployeeID]
  AND d.DeductionDate BETWEEN [pDateFrom] AND [pDateTo]
GROUP BY d.EmployeeID;
```

احفظه باسم: `qryEmployeeDeductionsByPeriod`

### 4.4 Employee cost by cost center

```sql
SELECT
  e.EmployeeID,
  e.FullName,
  c.CostCenterName,
  ec.AllocationPercent,
  ec.AllocatedAmount
FROM (tblEmployees AS e
INNER JOIN tblEmployeeCostCenter AS ec ON e.EmployeeID = ec.EmployeeID)
INNER JOIN tblCostCenters AS c ON ec.CostCenterID = c.CostCenterID
WHERE Nz(e.Status,'Active')='Active';
```

احفظه باسم: `qryEmployeeCostCenterDistribution`

### 4.5 Payroll calculator query (preview)

```sql
PARAMETERS pEmployeeID Long, pDateFrom DateTime, pDateTo DateTime;
SELECT
  e.EmployeeID,
  e.FullName,
  Nz(e.HourlyRate,0) AS HourlyRate,
  Nz(Sum(a.TotalHours),0) AS TotalWorkHours,
  Nz(Sum(a.OvertimeHours),0) AS TotalOvertimeHours,
  Nz(Sum(a.TotalHours),0) * Nz(e.HourlyRate,0) AS BaseSalary,
  Nz(Sum(a.OvertimeHours),0) * (Nz(e.HourlyRate,0) * 1.5) AS OvertimePay
FROM tblEmployees AS e
LEFT JOIN tblAttendance AS a ON e.EmployeeID = a.EmployeeID
WHERE e.EmployeeID=[pEmployeeID]
  AND a.WorkDate BETWEEN [pDateFrom] AND [pDateTo]
GROUP BY e.EmployeeID, e.FullName, e.HourlyRate;
```

احفظه باسم: `qryPayrollCalculationPreview`

---

## 5) Forms (Professional UX)

## Main Navigation Dashboard

- Form Name: `frmDashboard`
- Style: Modern flat theme (Navy/Slate + accent color)
- Buttons:
  - `btnEmployees` → `frmEmployees`
  - `btnAttendance` → `frmAttendanceEntry`
  - `btnWithdrawals` → `frmWithdrawals`
  - `btnPayroll` → `frmPayrollProcessor`
  - `btnCostCenters` → `frmCostCenters`
  - `btnReports` → `frmReportsHub`
  - `btnSearch` → `frmEmployeeSearch`
  - `btnLogout`

## Employees Management

- `frmEmployees`: CRUD كامل، حالة الموظف، الأجر بالساعة، بيانات التواصل.
- Subform: `sfrmEmployeeFinancialHistory` (آخر حضور/سحبيات/رواتب).

## Attendance Entry

- `frmAttendanceEntry`:
  - Combo `cboEmployeeID`
  - Date `txtWorkDate` default = Date()
  - `txtStartTime`, `txtEndTime`, `txtBreakMinutes`
  - `txtTotalHours` (locked)
  - `txtOvertimeHours` (auto: > 8 ساعات)
  - `btnSaveAttendance`

### Auto-calc expression (in form control source)

```vb
=Round(DateDiff("n",[StartTime],[EndTime])/60 - Nz([BreakMinutes],0)/60,2)
```

### Overtime expression

```vb
=IIf(Nz([txtTotalHours],0)>8, Nz([txtTotalHours],0)-8, 0)
```

## Withdrawals Entry

- `frmWithdrawals`:
  - موظف، تاريخ، مبلغ، نوع (Advance/Loan/Other)، وصف، اعتماد.

## Payroll Processing

- `frmPayrollProcessor`:
  - اختيار `Employee` أو All Employees
  - `PeriodStart`, `PeriodEnd`
  - زر `btnCalculatePayroll`
  - زر `btnPostPayroll`
  - Grid preview بالقيم المحسوبة

## Employee Search

- `frmEmployeeSearch`:
  - بحث بالاسم/الكود/الهاتف/الرقم الوطني
  - عرض فوري + فتح بطاقة الموظف

## Reports Hub

- `frmReportsHub`:
  - Daily Attendance
  - Detailed Payroll Slip
  - Total Withdrawals
  - Cost Center Distribution
  - Weekly Entitlements Cards
  - Full Employee Statement

---

## 6) VBA Automation (جاهز للنسخ)

أنشئ Module باسم `modPayroll`:

```vb
Option Compare Database
Option Explicit

Public Function CalculateEmployeePayroll(ByVal pEmployeeID As Long, ByVal pStart As Date, ByVal pEnd As Date) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlHours As String, sqlW As String, sqlD As String
    Dim totalHours As Double, otHours As Double
    Dim hourlyRate As Currency, baseSalary As Currency
    Dim otPay As Currency, withdrawals As Currency, deductions As Currency
    Dim netSalary As Currency

    Set db = CurrentDb

    hourlyRate = Nz(DLookup("HourlyRate", "tblEmployees", "EmployeeID=" & pEmployeeID), 0)

    sqlHours = "SELECT Nz(Sum(TotalHours),0) AS H, Nz(Sum(OvertimeHours),0) AS OTH " & _
               "FROM tblAttendance WHERE EmployeeID=" & pEmployeeID & _
               " AND WorkDate Between #" & Format(pStart, "yyyy-mm-dd") & "# And #" & Format(pEnd, "yyyy-mm-dd") & "#"

    Set rs = db.OpenRecordset(sqlHours)
    totalHours = Nz(rs!H, 0)
    otHours = Nz(rs!OTH, 0)
    rs.Close

    sqlW = "SELECT Nz(Sum(Amount),0) AS W FROM tblWithdrawals WHERE EmployeeID=" & pEmployeeID & _
           " AND WithdrawalDate Between #" & Format(pStart, "yyyy-mm-dd") & "# And #" & Format(pEnd, "yyyy-mm-dd") & "#"

    Set rs = db.OpenRecordset(sqlW)
    withdrawals = Nz(rs!W, 0)
    rs.Close

    sqlD = "SELECT Nz(Sum(Amount),0) AS D FROM tblDeductions WHERE EmployeeID=" & pEmployeeID & _
           " AND DeductionDate Between #" & Format(pStart, "yyyy-mm-dd") & "# And #" & Format(pEnd, "yyyy-mm-dd") & "#"

    Set rs = db.OpenRecordset(sqlD)
    deductions = Nz(rs!D, 0)
    rs.Close

    baseSalary = totalHours * hourlyRate
    otPay = otHours * (hourlyRate * 1.5)
    netSalary = baseSalary + otPay - withdrawals - deductions

    db.Execute "INSERT INTO tblPayroll (EmployeeID, PeriodStart, PeriodEnd, TotalWorkHours, TotalOvertimeHours, HourlyRate, BaseSalary, OvertimeRate, OvertimePay, TotalWithdrawals, TotalDeductions, NetSalary, PayrollStatus, CreatedAt) VALUES (" & _
               pEmployeeID & ", #" & Format(pStart, "yyyy-mm-dd") & "#, #" & Format(pEnd, "yyyy-mm-dd") & "#, " & _
               Round(totalHours,2) & ", " & Round(otHours,2) & ", " & hourlyRate & ", " & baseSalary & ", " & (hourlyRate * 1.5) & ", " & otPay & ", " & withdrawals & ", " & deductions & ", " & netSalary & ", 'Draft', Now())", dbFailOnError

    CalculateEmployeePayroll = True
    Exit Function

ErrHandler:
    CalculateEmployeePayroll = False
    MsgBox "Payroll calculation error: " & Err.Description, vbCritical
End Function
```

### Prevent duplicate attendance (Form Before Update)

```vb
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim n As Long

    n = DCount("*", "tblAttendance", "EmployeeID=" & Nz(Me.cboEmployeeID,0) & _
               " AND WorkDate=#" & Format(Me.txtWorkDate, "yyyy-mm-dd") & "#" & _
               " AND AttendanceID<>" & Nz(Me.AttendanceID,0))

    If n > 0 Then
        MsgBox "Attendance for this employee and date already exists.", vbExclamation
        Cancel = True
    End If
End Sub
```

### Smart validation before save

```vb
Private Sub btnSaveAttendance_Click()
    If IsNull(Me.cboEmployeeID) Then
        MsgBox "Please select employee.", vbExclamation: Exit Sub
    End If

    If IsNull(Me.txtWorkDate) Or IsNull(Me.txtStartTime) Or IsNull(Me.txtEndTime) Then
        MsgBox "Date, start time, and end time are required.", vbExclamation: Exit Sub
    End If

    If Me.txtEndTime <= Me.txtStartTime Then
        MsgBox "End time must be later than start time.", vbExclamation: Exit Sub
    End If

    DoCmd.RunCommand acCmdSaveRecord
    MsgBox "Attendance saved successfully.", vbInformation
End Sub
```

### Print official payment voucher

```vb
Public Sub PrintPaymentVoucher(ByVal pPayrollID As Long)
    DoCmd.OpenReport "rptPaymentVoucher", acViewPreview, , "PayrollID=" & pPayrollID
End Sub
```

---

## 7) Reports (Elegant + Practical)

- `rptDailyAttendance` (تاريخ يومي + حالة)
- `rptPayrollSlipDetailed` (كشف راتب تفصيلي لكل موظف)
- `rptTotalWithdrawals` (إجمالي سحبيات حسب الفترة)
- `rptCostCenterDistribution` (تحميل تكاليف حسب مركز)
- `rptWeeklyEntitlementCard` (كرت صرف أسبوعي)
- `rptEmployeeStatement` (ملف مالي كامل: حضور + إضافي + سحب + خصم + صافي)
- `rptConsolidatedPayroll` (كشف جماعي لكل الموظفين)

### PDF export (button on report form)

```vb
DoCmd.OutputTo acOutputReport, "rptPayrollSlipDetailed", acFormatPDF, CurrentProject.Path & "\\Exports\\Payroll_" & Format(Date, "yyyymmdd") & ".pdf", True
```

---

## 8) Security & Permissions

Role matrix in `tblUsers.UserRole`:

- `Admin`: Full access
- `Accountant`: Attendance, withdrawals, payroll, reports
- `Viewer`: Read-only reports/search

تطبيق الصلاحيات عند فتح النماذج:

- اخفاء/تعطيل أزرار الحذف والتعديل حسب الدور.
- منع فتح نماذج حساسة مباشرة بدون Session User.

> ملاحظة مهمة: لحماية قوية، يُفضّل تقسيم النظام إلى Front-End / Back-End مع صلاحيات ملفات NTFS أو Share permissions.

---

## 9) Audit Log (Financial traceability)

- سجّل كل تعديل/إضافة/حذف مهم في `tblAuditLog`.
- في `Form_BeforeUpdate` و `Form_AfterUpdate` استدعِ وظيفة `LogAudit`.

مخطط دالة:

```vb
Public Sub LogAudit(ByVal actionType As String, ByVal tableName As String, ByVal recordId As Long, ByVal fieldName As String, ByVal oldVal As String, ByVal newVal As String)
    CurrentDb.Execute "INSERT INTO tblAuditLog (UserID, ActionType, TableName, RecordID, FieldName, OldValue, NewValue, ActionTime) VALUES (" & Nz(TempVars!CurrentUserID,0) & ", '" & actionType & "', '" & tableName & "', " & recordId & ", '" & fieldName & "', '" & Replace(oldVal, "'", "''") & "', '" & Replace(newVal, "'", "''") & "', Now())"
End Sub
```

---

## 10) Backup Strategy

- عند إغلاق النظام أو عبر زر في `frmDashboard`:
  - نسخ ملف `.accdb` إلى مجلد `Backups` مع timestamp.
- احتفِظ بآخر 30 نسخة.

Pseudo VBA:

```vb
FileCopy CurrentProject.FullName, CurrentProject.Path & "\\Backups\\HRMS_" & Format(Now(), "yyyymmdd_hhnnss") & ".accdb"
```

---

## 11) UI/UX Standards (Elegant & Fast)

- خط واضح: Segoe UI / Tahoma.
- ألوان:
  - Primary: `#1F2A44`
  - Accent: `#1AA3A3`
  - Success: `#2E8B57`
  - Warning: `#E67E22`
- Action buttons with icons (save, print, search, payroll).
- Status strip in dashboard:
  - Logged-in user
  - Current date/time
  - Active payroll period

---

## 12) Recommended Additional Features (added)

- Multi-company readiness (`CompanyID` optional in core tables).
- Holiday calendar + attendance exception handling.
- Overtime policy tiers (1.25x / 1.5x / 2x).
- Late penalties auto-rule.
- Import attendance from Excel (`DoCmd.TransferSpreadsheet`).
- API-ready integration layer لاحقاً (if migrating to SQL Server).

---

## 13) Delivery Checklist (Go-Live)

1. إنشاء كل الجداول بالمفاتيح والمؤشرات.
2. ضبط العلاقات + RI + Cascade Update.
3. بناء النماذج: Dashboard / Employees / Attendance / Withdrawals / Payroll / Search / Reports.
4. إضافة أكواد VBA في Modules + Events.
5. تصميم التقارير وربطها بالاستعلامات.
6. اختبار سيناريو كامل:
   - موظف جديد → حضور أسبوع → سحبية → خصم → احتساب راتب → طباعة كشف → PDF.
7. تفعيل النسخ الاحتياطي.
8. مراجعة الصلاحيات.
9. مراجعة التدقيق المالي (Audit).

---

## 14) Suggested bilingual labels

- `frmEmployees.Caption = "إدارة الموظفين | Employees"`
- `frmPayrollProcessor.Caption = "احتساب الرواتب | Payroll Processing"`
- `rptPayrollSlipDetailed.Caption = "كشف راتب تفصيلي | Detailed Payroll Slip"`

هذا يعطيك نظام ثابت تقنياً باسماء إنجليزية، لكن واجهات وتقارير عربية واضحة وسهلة للمستخدم النهائي.
