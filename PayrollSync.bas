Option Explicit

' ============================================================
' PAYROLL SYNC - BAN MAP
' Chi can sua cot tai 2 ham:
'   PayrollColumnMap_ : sheet Salary
'   DSCNVColumnMap_   : sheet DSCNV
' Khong sua logic doc/JSON/API khi doi vi tri cot.
' Khoa nghiep vu: CitizenID 12 so + SalaryMonth.
' EmployeeCode chi la ma cua rieng tung thang.
' ============================================================

Private Const API_URL As String = _
    "https://script.google.com/macros/s/AKfycbzenFqwhLeihHWWp00noO5EK3g_buYizFpnOaNVqiVO2EdhtI_KO-RzEkSUahMst0Vapw/exec"

Private Const SYNC_API_KEY As String = _
    "TRUONGCONGVU_SALARYSLIP_1988_Elbalosnocni"

Private Const FILE_PASSWORD As String = "2410"

Private Const ROOT_PATH As String = _
    "\\192.168.0.253\vn hr\SALARY - 2014 - 2015\"

Private Const MAX_RETRY As Long = 3
Private Const WAIT_SECONDS As Long = 5
Private Const SHEET_START_ROW As Long = 7

Private Function PayrollColumnMap_() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare

    ' Salary: key logic -> column number
    m("employeeCode") = 4
    m("basicSalary") = 6
    m("workingDays") = 8
    m("holidayDays") = 9
    m("paidLeaveDays") = 10
    m("unpaidLeaveDays") = 11
    m("overtimeHours") = 12
    m("holidayWorkHours") = 13
    m("holidayOvertimeHours") = 14
    m("minimumRegionalLeaveDays") = 15
    m("nightHolidayOvertimeHours") = 16
    m("nightOvertimeHours") = 17
    m("holidayWorkDayHours") = 18
    m("holidayOvertimeDayHours") = 19
    m("nightHolidayOvertimeDayHours") = 20
    m("nightShiftDays") = 21
    m("otherMoney") = 24
    m("disciplinaryMoney") = 25
    m("loyalty2Years") = 26
    m("loyalty5Years") = 27
    m("loyalty10Years") = 28
    m("housingAllowance") = 29
    m("transportationAllowance") = 30
    m("attendanceBonus") = 31
    m("severanceAndUnusedLeave") = 32
    m("socialInsurance") = 33
    m("healthInsurance") = 34
    m("unemploymentInsurance") = 35
    m("otherDeductions") = 36
    m("advancePayment") = 37
    m("monthlySalary") = 40
    m("overtimeSalary") = 41
    m("commissionAndOverTargetBonus") = 43
    m("grossIncome") = 44
    m("personalIncomeTax") = 46
    m("netSalary") = 49
    m("fullName") = 84

    Set PayrollColumnMap_ = m
End Function

Private Function DSCNVColumnMap_() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare

    ' DSCNV: key logic -> column number
    m("fullName") = 2
    m("citizenID") = 8
    m("department") = 32       ' AF
    m("section") = 33          ' AG
    m("position") = 34         ' AH

    Set DSCNVColumnMap_ = m
End Function

Private Function MapCol_(ByVal m As Object, ByVal key As String) As Long
    If Not m.Exists(key) Then
        Err.Raise vbObjectError + 2101, "MapCol_", _
                  "Thieu map cot: " & key
    End If
    MapCol_ = CLng(m(key))
End Function

Public Sub RunPayrollSync()

    Dim targetDate As Date
    Dim targetMonth As Date
    Dim salaryMonth As String
    Dim snackFile As String
    Dim flexibleFile As String
    Dim snackRecords As Collection
    Dim flexibleRecords As Collection
    Dim allRecords As Collection
    Dim payload As String
    Dim responseText As String
    Dim success As Boolean

    On Error GoTo ErrorHandler

    LogMessage "=== BAT DAU DONG BO LUONG ==="
    LogMessage "Thoi gian chay: " & Format(Now, "dd/MM/yyyy HH:mm:ss")

    targetDate = Date
    targetMonth = DateAdd("m", -1, DateSerial(Year(targetDate), Month(targetDate), 1))
    salaryMonth = Format(targetMonth, "mm-yyyy")
    LogMessage "Thang luong can dong bo: " & salaryMonth

    snackFile = FindSnackFile_(targetMonth)
    flexibleFile = FindFlexibleFile_(targetMonth)

    Set allRecords = New Collection

    If Len(snackFile) > 0 Then
        Set snackRecords = ReadPayrollWorkbook_(snackFile, "Snack", targetMonth)
        AppendCollection_ allRecords, snackRecords
    Else
        LogMessage "Khong tim thay file Snack."
    End If

    If Len(flexibleFile) > 0 Then
        Set flexibleRecords = ReadPayrollWorkbook_(flexibleFile, "Flexible", targetMonth)
        AppendCollection_ allRecords, flexibleRecords
    Else
        LogMessage "Khong tim thay file Flexible."
    End If

    If allRecords.Count = 0 Then
        Err.Raise vbObjectError + 1001, "RunPayrollSync", _
                  "Khong co dong luong hop le (CCCD phai du 12 so)."
    End If

    LogMessage "Tong so dong du lieu hop le: " & CStr(allRecords.Count)

    payload = BuildSyncPayload_(allRecords, salaryMonth, snackFile, flexibleFile)
    LogMessage "Kich thuoc JSON: " & CStr(Len(payload)) & " ky tu."

    success = PostWithRetry_(payload, responseText)

    If Not success Then
        Err.Raise vbObjectError + 1002, "RunPayrollSync", _
                  "API dong bo that bai. Phan hoi: " & responseText
    End If

    LogMessage "Dong bo thanh cong."
    LogMessage "Phan hoi API: " & responseText
    LogMessage "=== KET THUC DONG BO ==="
    Exit Sub

ErrorHandler:
    LogMessage "LOI: " & Err.Number & " - " & Err.Description
    MsgBox "Dong bo luong that bai." & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Payroll Sync"
End Sub

Private Function FindSnackFile_(ByVal targetMonth As Date) As String
    Dim p As String
    p = ROOT_PATH & "VNLWW\" & Format(targetMonth, "yyyy") & "\SALARY " & _
        Format(targetMonth, "mm-yyyy") & ".xlsb"
    If FileExists_(p) Then FindSnackFile_ = p
End Function

Private Function FindFlexibleFile_(ByVal targetMonth As Date) As String
    Dim p As String
    p = ROOT_PATH & "Printing line\" & Format(targetMonth, "yyyy") & "\PRINTING LINE " & _
        Format(targetMonth, "mm-yyyy") & ".xlsb"
    If FileExists_(p) Then FindFlexibleFile_ = p
End Function

Private Function ReadPayrollWorkbook_( _
    ByVal filePath As String, _
    ByVal factoryName As String, _
    ByVal targetMonth As Date) As Collection

    Dim result As Collection
    Dim wb As Workbook
    Dim wsSalary As Worksheet
    Dim wsDSCNV As Worksheet
    Dim salaryMap As Object
    Dim dscnvMap As Object
    Dim rowIndex As Long
    Dim lastSalaryRow As Long
    Dim employeeName As String
    Dim normalizedName As String
    Dim employeeCode As String
    Dim citizenID As String
    Dim department As String
    Dim section As String
    Dim position As String
    Dim record As Object
    Dim stopReason As String

    On Error GoTo ErrorHandler
    Set result = New Collection
    Set salaryMap = PayrollColumnMap_()

    LogMessage "[" & factoryName & "] Dang mo file: " & filePath

    Set wb = Workbooks.Open(fileName:=filePath, password:=FILE_PASSWORD, _
                            ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)

    On Error Resume Next
    Set wsSalary = wb.Worksheets("Salary")
    Set wsDSCNV = wb.Worksheets("DSCNV")
    On Error GoTo ErrorHandler

    If wsSalary Is Nothing Then Err.Raise vbObjectError + 2001, "ReadPayrollWorkbook_", "Khong tim thay sheet Salary."
    If wsDSCNV Is Nothing Then Err.Raise vbObjectError + 2002, "ReadPayrollWorkbook_", "Khong tim thay sheet DSCNV."

    Set dscnvMap = BuildDSCNVMap_(wsDSCNV)
    lastSalaryRow = FindSalaryEndRow_(wsSalary, MapCol_(salaryMap, "fullName"))

    LogMessage "[" & factoryName & "] Dong doc: " & SHEET_START_ROW & " den " & lastSalaryRow

    For rowIndex = SHEET_START_ROW To lastSalaryRow

        If IsSalaryTotalRow_(wsSalary, rowIndex) Then
            stopReason = Trim$(CStr(wsSalary.Cells(rowIndex, 2).Value2))
            LogMessage "[" & factoryName & "] Gap dong TOTAL tai row " & rowIndex & " (" & stopReason & "), dung doc."
            Exit For
        End If

        employeeName = Trim$(CStr(wsSalary.Cells(rowIndex, MapCol_(salaryMap, "fullName")).Value2))
        If Len(employeeName) = 0 Then GoTo ContinueSalaryRow

        employeeCode = Trim$(CStr(wsSalary.Cells(rowIndex, MapCol_(salaryMap, "employeeCode")).Value2))
        normalizedName = NormalizeName_(employeeName)

        citizenID = ""
        department = ""
        section = ""
        position = ""

        If dscnvMap.Exists(normalizedName) Then
            citizenID = dscnvMap(normalizedName)("CitizenID")
            department = dscnvMap(normalizedName)("Department")
            section = dscnvMap(normalizedName)("Section")
            position = dscnvMap(normalizedName)("Position")
        End If

        ' Chi chap nhan dong co CCCD dung 12 chu so.
        If Not IsValidCitizenID_(citizenID) Then
            LogMessage "[" & factoryName & "] BO QUA row " & rowIndex & _
                       ": CCCD khong hop le - " & employeeName
            GoTo ContinueSalaryRow
        End If

        Set record = CreatePayrollRecord_(wsSalary, rowIndex, salaryMap)
        record("employeeCode") = employeeCode
        record("fullName") = employeeName
        record("citizenID") = citizenID
        record("department") = department
        record("section") = section
        record("position") = position

        result.Add record

ContinueSalaryRow:
    Next rowIndex

    wb.Close SaveChanges:=False
    Set wb = Nothing

    LogMessage "[" & factoryName & "] Doc duoc " & CStr(result.Count) & " dong hop le."
    Set ReadPayrollWorkbook_ = result
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Err.Raise Err.Number, "ReadPayrollWorkbook_", Err.Description
End Function

Private Function CreatePayrollRecord_( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal m As Object) As Object

    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r.CompareMode = vbTextCompare

    r("employeeCode") = ""
    r("fullName") = ""
    r("citizenID") = ""
    r("department") = ""
    r("section") = ""
    r("position") = ""

    r("basicSalary") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "basicSalary")))
    r("workingDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "workingDays")))
    r("holidayDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "holidayDays")))
    r("paidLeaveDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "paidLeaveDays")))
    r("unpaidLeaveDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "unpaidLeaveDays")))
    r("overtimeHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "overtimeHours")))
    r("holidayWorkHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "holidayWorkHours")))
    r("holidayOvertimeHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "holidayOvertimeHours")))
    r("minimumRegionalLeaveDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "minimumRegionalLeaveDays")))
    r("nightHolidayOvertimeHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "nightHolidayOvertimeHours")))
    r("nightOvertimeHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "nightOvertimeHours")))
    r("holidayWorkDayHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "holidayWorkDayHours")))
    r("holidayOvertimeDayHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "holidayOvertimeDayHours")))
    r("nightHolidayOvertimeDayHours") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "nightHolidayOvertimeDayHours")))
    r("nightShiftDays") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "nightShiftDays")))

    r("otherMoney") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "otherMoney")))
    r("disciplinaryMoney") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "disciplinaryMoney")))
    r("loyalty2Years") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "loyalty2Years")))
    r("loyalty5Years") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "loyalty5Years")))
    r("loyalty10Years") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "loyalty10Years")))
    r("housingAllowance") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "housingAllowance")))
    r("transportationAllowance") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "transportationAllowance")))
    r("attendanceBonus") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "attendanceBonus")))
    r("severanceAndUnusedLeave") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "severanceAndUnusedLeave")))

    r("monthlySalary") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "monthlySalary")))
    r("overtimeSalary") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "overtimeSalary")))
    r("commissionAndOverTargetBonus") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "commissionAndOverTargetBonus")))

    ' III - tong cac khoan phu cap/tro cap/ho tro/bo sung khac.
    r("otherIncome") = _
        r("otherMoney") + r("disciplinaryMoney") + _
        r("loyalty2Years") + r("loyalty5Years") + r("loyalty10Years") + _
        r("housingAllowance") + r("transportationAllowance") + _
        r("attendanceBonus") + r("commissionAndOverTargetBonus") + _
        r("severanceAndUnusedLeave")

    r("otherDeductions") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "otherDeductions")))
    r("advancePayment") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "advancePayment")))
    r("socialInsurance") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "socialInsurance")))
    r("healthInsurance") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "healthInsurance")))
    r("unemploymentInsurance") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "unemploymentInsurance")))
    r("personalIncomeTax") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "personalIncomeTax")))
    r("grossIncome") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "grossIncome")))
    r("netSalary") = CellNumber_(ws.Cells(rowIndex, MapCol_(m, "netSalary")))

    Set CreatePayrollRecord_ = r
End Function

Private Function BuildDSCNVMap_(ByVal ws As Worksheet) As Object
    Dim map As Object, item As Object, m As Object
    Dim lastRow As Long, rowIndex As Long
    Dim employeeName As String, key As String, citizenID As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    Set m = DSCNVColumnMap_()

    lastRow = LastUsedRow_(ws, MapCol_(m, "fullName"))

    For rowIndex = SHEET_START_ROW To lastRow
        employeeName = Trim$(CStr(ws.Cells(rowIndex, MapCol_(m, "fullName")).Value2))
        If Len(employeeName) > 0 Then
            citizenID = ReadCitizenID_(ws.Cells(rowIndex, MapCol_(m, "citizenID")))
            If IsValidCitizenID_(citizenID) Then
                key = NormalizeName_(employeeName)
                Set item = CreateObject("Scripting.Dictionary")
                item("CitizenID") = citizenID
                item("Department") = Trim$(CStr(ws.Cells(rowIndex, MapCol_(m, "department")).Value2))
                item("Section") = Trim$(CStr(ws.Cells(rowIndex, MapCol_(m, "section")).Value2))
                item("Position") = Trim$(CStr(ws.Cells(rowIndex, MapCol_(m, "position")).Value2))

                If map.Exists(key) Then
                    LogMessage "CANH BAO: Trung ten DSCNV, uu tien dong dau: " & employeeName
                Else
                    map.Add key, item
                End If
            End If
        End If
    Next rowIndex

    Set BuildDSCNVMap_ = map
End Function

Private Function IsSalaryTotalRow_(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    Dim v As String
    v = UCase$(Trim$(CStr(ws.Cells(rowIndex, 2).Value2)))
    IsSalaryTotalRow_ = (Left$(v, 5) = "TOTAL")
End Function

Private Function FindSalaryEndRow_(ByVal ws As Worksheet, ByVal nameCol As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = LastUsedRow_(ws, nameCol)
    For r = SHEET_START_ROW To lastRow
        If IsSalaryTotalRow_(ws, r) Then
            FindSalaryEndRow_ = r
            Exit Function
        End If
    Next r
    FindSalaryEndRow_ = lastRow
End Function

Private Function IsValidCitizenID_(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    value = Replace(Replace(Trim$(value), "'", ""), " ", "")
    If Len(value) <> 12 Then Exit Function
    For i = 1 To 12
        ch = Mid$(value, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsValidCitizenID_ = True
End Function


Private Function ReadCitizenID_(ByVal cell As Range) As String
    Dim textValue As String
    Dim numberValue As Double
    Dim formatText As String

    On Error GoTo Fallback

    textValue = Trim$(CStr(cell.Value2))
    textValue = Replace(textValue, "'", "")
    textValue = Replace(textValue, " ", "")

    If Len(textValue) = 0 Then Exit Function

    If IsNumeric(cell.Value2) Then
        numberValue = CDbl(cell.Value2)
        formatText = CStr(cell.NumberFormat)

        If Len(formatText) > 0 And InStr(1, formatText, "0", vbTextCompare) > 0 Then
            On Error Resume Next
            textValue = Format$(numberValue, Replace(Replace(formatText, """", ""), " ", ""))
            On Error GoTo Fallback
        Else
            textValue = Format$(numberValue, "0")
        End If
    End If

    textValue = Replace(textValue, ".", "")
    textValue = Replace(textValue, ",", "")

    ReadCitizenID_ = textValue
    Exit Function

Fallback:
    ReadCitizenID_ = Trim$(Replace(CStr(cell.Text), "'", ""))
End Function

Private Function NormalizeName_( _
    ByVal value As String) As String

    Dim textValue As String

    textValue = _
        Trim$(value)

    textValue = _
        Replace(textValue, vbTab, " ")

    Do While InStr(textValue, "  ") > 0

        textValue = _
            Replace( _
                textValue, _
                "  ", _
                " ")

    Loop

    NormalizeName_ = _
        LCase$(textValue)

End Function

Private Function CellNumber_( _
    ByVal cell As Range) As Double

    Dim value As Variant
    Dim textValue As String

    On Error GoTo SafeZero

    value = cell.Value2

    If IsNumeric(value) Then
        CellNumber_ = CDbl(value)
        Exit Function
    End If

    textValue = _
        Trim$(CStr(value))

    textValue = _
        Replace(textValue, ",", "")

    If IsNumeric(textValue) Then
        CellNumber_ = CDbl(textValue)
    Else
        CellNumber_ = 0
    End If

    Exit Function

SafeZero:

    CellNumber_ = 0

End Function

Private Function LastUsedRow_( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long) As Long

    LastUsedRow_ = _
        ws.Cells( _
            ws.Rows.count, _
            columnNumber).End(xlUp).row

End Function

Private Function FileExists_( _
    ByVal filePath As String) As Boolean

    FileExists_ = _
        (Len(Dir$(filePath, vbNormal)) > 0)

End Function

Private Sub AppendCollection_( _
    ByVal target As Collection, _
    ByVal source As Collection)

    Dim item As Variant

    If source Is Nothing Then
        Exit Sub
    End If

    For Each item In source
        target.Add item
    Next item

End Sub

Private Function BuildSyncPayload_( _
    ByVal records As Collection, _
    ByVal salaryMonth As String, _
    ByVal snackFile As String, _
    ByVal flexibleFile As String) As String

    Dim json As String
    Dim i As Long

    json = "{"

    json = json & _
        """action"":""syncPayroll"","

    json = json & _
        """apiKey"":""" & _
        JsonEscape_(SYNC_API_KEY) & _
        ""","

    json = json & _
        """salaryMonth"":""" & _
        JsonEscape_(salaryMonth) & _
        ""","

    json = json & _
        """source"":{"

    json = json & _
        """snackFile"":""" & _
        JsonEscape_(snackFile) & _
        ""","

    json = json & _
        """flexibleFile"":""" & _
        JsonEscape_(flexibleFile) & _
        """},"

    json = json & _
        """records"":["

    For i = 1 To records.count

        If i > 1 Then
            json = json & ","
        End If

        json = json & _
            RecordToJson_(records(i))

    Next i

    json = json & "]"

    json = json & "}"

    BuildSyncPayload_ = json

End Function

Private Function RecordToJson_( _
    ByVal record As Object) As String

    Dim json As String

    json = "{"

    AddJsonString_ json, _
        "employeeCode", _
        record("employeeCode"), True

    AddJsonString_ json, _
        "fullName", _
        record("fullName"), True

    AddJsonString_ json, _
        "citizenID", _
        record("citizenID"), True

    AddJsonString_ json, _
        "department", _
        record("department"), True

    AddJsonString_ json, _
        "section", _
        record("section"), True

    AddJsonString_ json, _
        "position", _
        record("position"), True

    AddJsonNumber_ json, _
        "basicSalary", _
        record("basicSalary"), True

    AddJsonNumber_ json, _
        "workingDays", _
        record("workingDays"), True

    AddJsonNumber_ json, _
        "holidayDays", _
        record("holidayDays"), True

    AddJsonNumber_ json, _
        "paidLeaveDays", _
        record("paidLeaveDays"), True

    AddJsonNumber_ json, _
        "unpaidLeaveDays", _
        record("unpaidLeaveDays"), True

    AddJsonNumber_ json, _
        "overtimeHours", _
        record("overtimeHours"), True

    AddJsonNumber_ json, _
        "holidayWorkHours", _
        record("holidayWorkHours"), True

    AddJsonNumber_ json, _
        "holidayOvertimeHours", _
        record("holidayOvertimeHours"), True

    AddJsonNumber_ json, _
        "minimumRegionalLeaveDays", _
        record("minimumRegionalLeaveDays"), True

    AddJsonNumber_ json, _
        "nightHolidayOvertimeHours", _
        record("nightHolidayOvertimeHours"), True

    AddJsonNumber_ json, _
        "nightOvertimeHours", _
        record("nightOvertimeHours"), True

    AddJsonNumber_ json, _
        "holidayWorkDayHours", _
        record("holidayWorkDayHours"), True

    AddJsonNumber_ json, _
        "holidayOvertimeDayHours", _
        record("holidayOvertimeDayHours"), True

    AddJsonNumber_ json, _
        "nightShiftDays", _
        record("nightShiftDays"), True

    AddJsonNumber_ json, _
        "nightHolidayOvertimeDayHours", _
        record("nightHolidayOvertimeDayHours"), True

    AddJsonNumber_ json, _
        "otherMoney", _
        record("otherMoney"), True

    AddJsonNumber_ json, _
        "disciplinaryMoney", _
        record("disciplinaryMoney"), True

    AddJsonNumber_ json, _
        "loyalty2Years", _
        record("loyalty2Years"), True

    AddJsonNumber_ json, _
        "loyalty5Years", _
        record("loyalty5Years"), True

    AddJsonNumber_ json, _
        "loyalty10Years", _
        record("loyalty10Years"), True

    AddJsonNumber_ json, _
        "housingAllowance", _
        record("housingAllowance"), True

    AddJsonNumber_ json, _
        "transportationAllowance", _
        record("transportationAllowance"), True

    AddJsonNumber_ json, _
        "attendanceBonus", _
        record("attendanceBonus"), True

    AddJsonNumber_ json, _
        "severanceAndUnusedLeave", _
        record("severanceAndUnusedLeave"), True

    AddJsonNumber_ json, _
        "monthlySalary", _
        record("monthlySalary"), True

    AddJsonNumber_ json, _
        "overtimeSalary", _
        record("overtimeSalary"), True

    AddJsonNumber_ json, _
        "commissionAndOverTargetBonus", _
        record("commissionAndOverTargetBonus"), True

    AddJsonNumber_ json, _
        "otherIncome", _
        record("otherIncome"), True

    AddJsonNumber_ json, _
        "otherDeductions", _
        record("otherDeductions"), True

    AddJsonNumber_ json, _
        "advancePayment", _
        record("advancePayment"), True

    AddJsonNumber_ json, _
        "socialInsurance", _
        record("socialInsurance"), True

    AddJsonNumber_ json, _
        "healthInsurance", _
        record("healthInsurance"), True

    AddJsonNumber_ json, _
        "unemploymentInsurance", _
        record("unemploymentInsurance"), True

    AddJsonNumber_ json, _
        "personalIncomeTax", _
        record("personalIncomeTax"), True

    AddJsonNumber_ json, _
            "grossIncome", _
        record("grossIncome"), True
            
    AddJsonNumber_ json, _
        "netSalary", _
        record("netSalary"), False

    json = json & "}"

    RecordToJson_ = json

End Function

Private Sub AddJsonString_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As String, _
    ByVal addComma As Boolean)

    json = json & _
        """" & key & """:""" & _
        JsonEscape_(value) & """"

    If addComma Then
        json = json & ","
    End If

End Sub

Private Sub AddJsonNumber_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As Double, _
    ByVal addComma As Boolean)

    ' IMPORTANT: payroll numeric values are sent as JSON strings.
    ' This completely avoids locale-dependent VBA number serialization
    ' such as 0., .5, or 1,5 which are invalid JSON in some cases.
    ' Code.gs converts these strings back to real numbers before writing
    ' them into Google Sheets.
    json = json & _
        """" & key & """:""" & _
        JsonNumberString_(value) & """"

    If addComma Then
        json = json & ","
    End If

End Sub

Private Function JsonNumberString_( _
    ByVal value As Double) As String

    Dim result As String

    If IsNumeric(value) Then
        result = Format$(value, "0.############################")
    Else
        result = "0"
    End If

    ' Force JSON-independent decimal point. Format$ above always supplies
    ' the leading zero because the integer part is mandatory in the format.
    If Application.International(xlDecimalSeparator) <> "." Then
        result = Replace$(result, _
            Application.International(xlDecimalSeparator), ".")
    End If

    If Len(result) = 0 Or result = "-" Or result = "." Or result = "-." Then
        result = "0"
    End If

    If Left$(result, 1) = "." Then
        result = "0" & result
    ElseIf Left$(result, 2) = "-." Then
        result = "-0" & Mid$(result, 2)
    End If

    JsonNumberString_ = result

End Function

Private Function JsonEscape_( _
    ByVal value As String) As String

    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String

    result = ""

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)

        If code < 0 Then code = code + 65536

        Select Case code
            Case 8
                result = result & "\b"
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 12
                result = result & "\f"
            Case 13
                result = result & "\r"
            Case 0 To 7, 11, 14 To 31
                result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case 34
                result = result & "\"""
            Case 92
                result = result & "\\"
            Case Else
                ' Encode non-ASCII UTF-16 code units as JSON \uXXXX.
                ' This keeps the entire request ASCII and avoids Excel/VBA
                ' code-page/UTF-8 conversion problems with Vietnamese text.
                If code > 127 Then
                    result = result & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i

    JsonEscape_ = result

End Function

Private Function PostWithRetry_( _
    ByVal payload As String, _
    ByRef responseText As String) As Boolean

    Dim attempt As Long
    Dim waitSeconds As Long

    For attempt = 1 To MAX_RETRY

        LogMessage _
            "API attempt " & _
            CStr(attempt) & _
            "/" & _
            CStr(MAX_RETRY)

        responseText = _
            HttpPostJson_(payload)

        LogMessage _
            "API response: " & _
            Left$(responseText, 1000)

        If ResponseOK_(responseText) Then

            PostWithRetry_ = True

            Exit Function

        End If

        If attempt < MAX_RETRY Then

            waitSeconds = _
                WAIT_SECONDS * attempt

            LogMessage _
                "API loi. Cho " & _
                CStr(waitSeconds) & _
                " giay roi thu lai."

            SleepSeconds_ waitSeconds

        End If

    Next attempt

    PostWithRetry_ = False

End Function

Private Function HttpPostJson_( _
    ByVal payload As String) As String

    Dim http As Object
    Dim bodyBytes As Variant

    On Error GoTo ErrorHandler

    Set http = _
        CreateObject( _
            "WinHttp.WinHttpRequest.5.1")

    bodyBytes = Utf8Bytes_(payload)

    http.Open _
        "POST", _
        API_URL, _
        False

    http.SetTimeouts _
        15000, _
        15000, _
        30000, _
        30000

    http.SetRequestHeader _
        "Content-Type", _
        "application/json; charset=utf-8"

    http.Send bodyBytes

    HttpPostJson_ = _
        CStr(http.responseText)

    Exit Function

ErrorHandler:

    HttpPostJson_ = _
        "{""success"":false,""error"":""" & _
        JsonEscape_(Err.Description) & _
        """}"

End Function

Private Function Utf8Bytes_( _
    ByVal textValue As String) As Variant

    Dim stream As Object
    Dim bytes As Variant
    Dim length As Long

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.position = 0
    stream.Type = 1

    bytes = stream.Read
    length = UBound(bytes) - LBound(bytes) + 1

    ' ADODB.Stream writes UTF-8 BOM (EF BB BF).
    If length >= 3 Then
        If bytes(0) = &HEF And _
           bytes(1) = &HBB And _
           bytes(2) = &HBF Then

            stream.position = 3
            bytes = stream.Read
        End If
    End If

    stream.Close
    Set stream = Nothing

    Utf8Bytes_ = bytes

End Function

Private Function ResponseOK_( _
    ByVal responseText As String) As Boolean

    Dim textValue As String

    textValue = _
        LCase$(Trim$(responseText))

    If InStr(1, textValue, _
             """success"":true", _
             vbTextCompare) > 0 Then

        ResponseOK_ = True

    Else

        ResponseOK_ = False

    End If

End Function

Private Sub SleepSeconds_( _
    ByVal seconds As Long)

    Dim endTime As Date

    endTime = _
        DateAdd( _
            "s", _
            seconds, _
            Now)

    Do While Now < endTime
        DoEvents
    Loop

End Sub

Private Sub LogMessage( _
    ByVal message As String)

    Dim logFile As String
    Dim fileNumber As Integer

    logFile = _
        ThisWorkbook.path & _
        "\PayrollSync.log"

    fileNumber = _
        FreeFile

    Open logFile _
        For Append _
        As #fileNumber

    Print #fileNumber, _
        Format(Now, "yyyy-mm-dd HH:mm:ss") & _
        " | " & _
        message

    Close #fileNumber

End Sub

