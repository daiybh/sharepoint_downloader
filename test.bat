@echo off
powershell.exe -ExecutionPolicy Bypass -File "Download-SharePointFile.ps1"

if %ERRORLEVEL% equ 0 (
    echo success
) else (
    echo failed, code: %ERRORLEVEL%
)