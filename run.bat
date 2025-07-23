@echo off
setlocal

:: Define the full path to the Excel macro file
set FILE_PATH=%~dp0macros\Dashboard of Operations.xlsm

:: Call the VBScript to open Excel and run the macro
cscript //nologo script.vbs "%FILE_PATH%"

endlocal
