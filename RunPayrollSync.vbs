Option Explicit

Dim excelApp
Dim workbook
Dim workbookPath
Dim fso
Dim scriptFolder

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
workbookPath = fso.BuildPath(scriptFolder, "PayrollSync.xlsm")

On Error Resume Next

Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Khong the khoi dong Microsoft Excel: " & Err.Description
    WScript.Quit 1
End If
Err.Clear

excelApp.Visible = False
excelApp.DisplayAlerts = False
excelApp.EnableEvents = False
excelApp.ScreenUpdating = False

Set workbook = excelApp.Workbooks.Open(workbookPath, 0, False)
If Err.Number <> 0 Then
    WScript.Echo "Khong the mo file: " & workbookPath & vbCrLf & Err.Description
    excelApp.Quit
    WScript.Quit 1
End If
Err.Clear

excelApp.Run "'" & workbook.Name & "'!RunPayrollSync"

If Err.Number <> 0 Then
    WScript.Echo "Macro dong bo gap loi: " & Err.Description
    Err.Clear
End If

If Not workbook Is Nothing Then workbook.Close False
If Not excelApp Is Nothing Then excelApp.Quit

Set workbook = Nothing
Set excelApp = Nothing
Set fso = Nothing

WScript.Quit 0
