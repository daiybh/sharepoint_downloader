@echo off
powershell.exe -ExecutionPolicy Bypass -File "Download-SharePointFile.ps1" -SharePointURL "https://riedelcommunications.sharepoint.com/:u:/r/sites/SimplyLiveInternal/Shared Documents/R&D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip?csf=1&web=1&e=MnZxu8" -SaveFileName "TcTableAnalyzer.zip"

if %ERRORLEVEL% equ 0 (
    echo success
) else (
    echo failed, code: %ERRORLEVEL%
)