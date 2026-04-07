@echo off


set link="https://riedelcommunications.sharepoint.com/:u:/r/sites/SimplyLiveInternal/Shared%20Documents/R%26D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip?csf=1&web=1&e=StSWpS"
set savepath="TcTableAnalyzer11.26.4.5.zip"
echo using default link and save path

call download.bat %link% %savepath%    

@REM set dropboxLink="https://www.dropbox.com/scl/fi/bs0ars5nx5rwlca0jrizv/WebConfig-1.25.52.231.exe?rlkey=sfo8d4d3pa6xom4m472i9i5cz&dl=1"
@REM set target_version=1.25.52.231
@REM set filePath="%~dp0Dependent\WebConfig.exe"
@REM call :CheckVersion %filePath%
@REM if %ERRORLEVEL% neq 0 (    
@REM     call download.bat %dropboxLink% %filePath%    
@REM     @REM if %ERRORLEVEL% neq 0 (        
@REM     @REM     exit /b 1
@REM     @REM )
@REM )