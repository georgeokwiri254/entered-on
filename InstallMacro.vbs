' VBScript to install the macro into the Excel file
' Double-click this file to automatically install the macro

Option Explicit

Dim objExcel, objWorkbook, objModule
Dim strExcelPath, strVBAPath, strVBACode
Dim objFSO, objFile
Dim WshShell

' Create objects
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get current folder
Dim currentFolder
currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Set paths
strExcelPath = currentFolder & "\22-11-2025 Entered On.xlsm"
strVBAPath = currentFolder & "\ProcessReservations.vba"

' Check if files exist
If Not objFSO.FileExists(strExcelPath) Then
    MsgBox "Error: Excel file not found at:" & vbCrLf & strExcelPath, vbCritical, "File Not Found"
    WScript.Quit
End If

If Not objFSO.FileExists(strVBAPath) Then
    MsgBox "Error: VBA file not found at:" & vbCrLf & strVBAPath, vbCritical, "File Not Found"
    WScript.Quit
End If

' Read VBA code
Set objFile = objFSO.OpenTextFile(strVBAPath, 1)
strVBACode = objFile.ReadAll
objFile.Close

' Create Excel object
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Error: Could not create Excel application." & vbCrLf & "Error: " & Err.Description, vbCritical, "Excel Error"
    WScript.Quit
End If
On Error GoTo 0

' Make Excel visible (optional)
objExcel.Visible = False
objExcel.DisplayAlerts = False

' Open workbook
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)
If Err.Number <> 0 Then
    MsgBox "Error: Could not open workbook." & vbCrLf & "Error: " & Err.Description, vbCritical, "Excel Error"
    objExcel.Quit
    WScript.Quit
End If
On Error GoTo 0

' Add module and code
On Error Resume Next
Set objModule = objWorkbook.VBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
If Err.Number <> 0 Then
    MsgBox "Error: Could not add VBA module." & vbCrLf & vbCrLf & _
           "This might be due to macro security settings." & vbCrLf & _
           "Please enable 'Trust access to the VBA project object model' in:" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings", _
           vbCritical, "VBA Access Error"
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit
End If

objModule.Name = "ReservationProcessor"
objModule.CodeModule.AddFromString strVBACode
On Error GoTo 0

' Save and close
objWorkbook.Save
objWorkbook.Close

' Clean up
objExcel.Quit
Set objModule = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing
Set WshShell = Nothing

MsgBox "Macro successfully installed in:" & vbCrLf & strExcelPath & vbCrLf & vbCrLf & _
       "You can now open the file and run the macro by pressing ALT+F8 and selecting 'ProcessEnteredOnReport'", _
       vbInformation, "Installation Complete"
