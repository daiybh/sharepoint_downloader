@echo off

set link="%~1"
set savepath="%~2"
echo link: %link%
echo save path: %savepath%

REM check if link is from sharepoint
echo %link% | findstr /i "sharepoint" >nul
if %ERRORLEVEL% equ 0 (
    powershell.exe -ExecutionPolicy Bypass -File "Download-SharePointFile.ps1" -SharePointURL %link% -SaveFileName %savepath%
) else (
    REM check if link is from dropbox
    "%~dp0curl.exe" -L -o %savePath% %link%
)

if %ERRORLEVEL% equ 0 (
    echo success
) else (
    echo failed, code: %ERRORLEVEL%
)