  set localFile="D:\Codes\Simplylive\sharepoint_downloader\downloadUrl_TcTableAnalyzer11.26.4.5.zip"
  set remotePath="dd/12.12.12/dd.zip"
  powershell.exe -ExecutionPolicy Bypass -File "Download-SharePointFile.ps1" -UploadLocalFile %localFile% -UploadDestPath %remotePath%