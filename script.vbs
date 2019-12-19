Dim args, objExcel

set args = wscript.arguments
set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open args(0)
ObjExcel.Visible = True
ObjExcel.Run "refresh"
ObjExcel.Quit

