' script.vbs - Launches Excel, opens a workbook, and runs the "refresh" macro

Dim args, objExcel, workbookPath

Set args = WScript.Arguments

If args.Count = 0 Then
    WScript.Echo "❌ Error: No Excel file path provided."
    WScript.Quit 1
End If

workbookPath = args(0)

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

On Error Resume Next
objExcel.Workbooks.Open workbookPath

If Err.Number <> 0 Then
    WScript.Echo "❌ Failed to open workbook: " & workbookPath
    objExcel.Quit
    WScript.Quit 2
End If

objExcel.Run "refresh"
objExcel.Quit
