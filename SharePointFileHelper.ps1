param(
    [Parameter(Mandatory=$false)][string]$SharePointURL,
    [Parameter(Mandatory=$false)][string]$SaveFileName = "",
    [Parameter(Mandatory=$false)][string]$UploadLocalFile = "",
    [Parameter(Mandatory=$false)][string]$UploadDestPath = "",
    [Parameter(Mandatory=$false)][bool]$EnableLogging = $true,
    [Parameter(Mandatory=$false)][string]$SaveDir = ".",
    [Parameter(Mandatory=$false)][string]$LogFile = "sharepoint_downloader.log",
    [Parameter(Mandatory=$false)][string]$ConfigFile = "config.json"
)

# 全局变量
$script:LogFile = $LogFile
# 设置日志
function Setup-Logger {
    if ($EnableLogging) {     
        if (-not (Test-Path -Path $script:LogFile)) {
            New-Item -ItemType File -Path $script:LogFile | Out-Null
        }
    }
}

# 日志输出到文件和控制台
function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][ValidateSet("INFO", "ERROR", "WARNING")][string]$Level = "INFO"
    )
    
    if (-not $EnableLogging) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp $Level $Message"
    
    Write-Host $logMessage
    Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8
}

# 解析 SharePoint URL
function Split-SharePointURL {
    param([string]$SharedURL)
    
    try {
        $uri = [System.Uri]$SharedURL

        
        $domain = $uri.Host
        # because the sharedpointURL is  R%26D , but in bat the %2 was encoded
        $absolutePath= $uri.AbsolutePath.Trim('/') -replace '/R6D/', '/R&D/' 
        $absolutePath = $absolutePath -replace '/Shared0Documents/', '/Shared Documents/'
        $absolutePath = $absolutePath -replace '/Shared%20Documents/', '/Shared Documents/'

        $path = $absolutePath -split '/'
        
        $sitesIndex = $path.IndexOf('sites')
        $docsIndex = $path.IndexOf('Shared Documents')

        if ($sitesIndex -eq -1 -or $docsIndex -eq -1) {
            Write-Log "URL format is incorrect, unable to find 'sites' or 'Shared Documents'" "ERROR"
            return $null
        }
        
        $sitePath = ($path[$sitesIndex..($sitesIndex + 1)] -join '/')
        $filePath = ($path[($docsIndex+1)..($path.Length-1)] -join '/')

        return @{
            Domain = $domain
            SitePath = $sitePath
            FilePath = $filePath
        }
    }
    catch {
        Write-Log "Failed to parse URL: $_" "ERROR"
        return $null
    }
}

# 获取 SiteID
function Get-SiteID {
    param(
        [string]$Domain,
        [string]$SitePath,
        [string]$Token
    )
    
    $url = "https://graph.microsoft.com/v1.0/sites/${Domain}:/${SitePath}"
    Write-Log "####################"
    Write-Log "get_site_id"
    Write-Log $url
    
    try {
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        return $response.id
    }
    catch {
        Write-Log "Failed to get siteID: $_" "ERROR"
        return $null
    }
}

# 获取 DriveID
function Get-DriveID {
    param(
        [string]$SiteID,
        [string]$Token
    )
    
    $url = "https://graph.microsoft.com/v1.0/sites/${SiteID}/drives"
    Write-Log "####################"
    Write-Log "get_drive_id"
    Write-Log $url
    
    try {
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        if ($response.value -and $response.value.Count -gt 0) {
            return $response.value[0].id
        }
        
        Write-Log "No drives found for site" "ERROR"
        return $null
    }
    catch {
        Write-Log "Failed to get driveID: $_" "ERROR"
        return $null
    }
}

# 获取文件信息
function Get-FileInfo {
    param(
        [string]$SiteID,
        [string]$DriveID,
        [string]$FilePath,
        [string]$Token
    )
    
    $url = "https://graph.microsoft.com/v1.0/sites/${SiteID}/drives/${DriveID}/root:/${FilePath}"
    Write-Log "####################"
    Write-Log "get_file_info"
    Write-Log $url
    
    try {
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Log "Failed to get file info: $_" "ERROR"
        return $null
    }
}

# 下载文件
function Download-File {
    param(
        [string]$DownloadURL,
        [string]$LocalPath
    )
    
    try {
        Write-Log "Downloading $LocalPath from URL: $DownloadURL"
        
        Invoke-WebRequest -Uri $DownloadURL -OutFile $LocalPath -UseBasicParsing
        
        Write-Log "Download completed: $LocalPath"
        return $true
    }
    catch {
        Write-Log "Download failed: $_" "ERROR"
        return $false
    }
}

# 从 SharePoint 下载
function Download-FromSharePoint {
    param(
        [string]$SharedURL,
        [string]$SaveDir,
        [string]$Token
    )
    
    # 解析 URL
    $urlParts = Split-SharePointURL -SharedURL $SharedURL
    if (-not $urlParts) {
        Write-Log "URL parsing failed, aborting download." "ERROR"
        return $false
    }
    
    $domain = $urlParts.Domain
    $sitePath = $urlParts.SitePath
    $filePath = $urlParts.FilePath
    Write-Log "Domain: $domain, SitePath: $sitePath, FilePath: $filePath"
    
    # 获取 SiteID
    $siteID = Get-SiteID -Domain $domain -SitePath $sitePath -Token $Token
    if (-not $siteID) {
        Write-Log "Unable to get siteID, aborting download." "ERROR"
        return $false
    }
    Write-Log "SiteID: $siteID"
    
    # 获取 DriveID
    $driveID = Get-DriveID -SiteID $siteID -Token $Token
    if (-not $driveID) {
        Write-Log "Unable to get driveID, aborting download." "ERROR"
        return $false
    }
    Write-Log "DriveID: $driveID"
    
    # 获取文件信息
    $fileInfo = Get-FileInfo -SiteID $siteID -DriveID $driveID -FilePath $filePath -Token $Token
    if (-not $fileInfo -or -not $fileInfo.'@microsoft.graph.downloadUrl') {
        Write-Log "Unable to get download link, aborting download." "ERROR"
        return $false
    }
    
    $downloadURL = $fileInfo.'@microsoft.graph.downloadUrl'
    Write-Log "@microsoft.graph.downloadUrl $downloadURL"
    
    # 准备本地路径
    $fileName = Split-Path -Leaf $filePath
    $localPath = Join-Path $SaveDir "$fileName"
    if ($SaveFileName) {
        $localPath = Join-Path $SaveDir $SaveFileName
    }
    # 下载文件
    return Download-File -DownloadURL $downloadURL -LocalPath $localPath
}

function LoadEnv {
    param (
        [string]$Path = ".env"
    )
    # 读取文件并遍历每一行
    Get-Content $Path | ForEach-Object {
        # 忽略空行和注释行
        if ($_ -and $_ -notmatch '^\s*#') {
            # 按第一个等号分割键值对
            $key, $value = $_ -split '=', 2
            # 设置环境变量
            Set-Item -Path "env:$key" -Value $value
        }
    }    
}

# 读取配置文件
function Load-Config {
    param (
        [string]$ConfigPath
    )
    
    if (-not (Test-Path -Path $ConfigPath)) {
        Write-Log "Config file not found: $ConfigPath" "WARNING"
        return $null
    }
    
    try {
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        Write-Log "Config file loaded successfully"
        return $config
    }
    catch {
        Write-Log "Failed to load config file: $_" "ERROR"
        return $null
    }
}

function Upload-LargeFileToSharePoint {
    param(
        [string]$SiteID,
        [string]$DriveID,
        [string]$FilePath,        # SharePoint 中的目标路径
        [string]$LocalFilePath,   # 本地文件路径
        [string]$Token,
        [int]$ChunkSize = 5 * 1024 * 1024  # 默认 5MB 每片
    )
    
    $uploadUrl = "https://graph.microsoft.com/v1.0/sites/${SiteID}/drives/${DriveID}/root:/${FilePath}:/createUploadSession"
    Write-Log "####################"
    Write-Log "upload_large_file_to_sharepoint"
    Write-Log "Target URL: $uploadUrl"
    Write-Log "Local file: $LocalFilePath"
    
    try {
        # 检查本地文件
        if (-not (Test-Path $LocalFilePath)) {
            Write-Log "Local file not found: $LocalFilePath" "ERROR"
            return $null
        }
        
        $fileInfo = Get-Item $LocalFilePath
        $fileSize = $fileInfo.Length
        Write-Log "File size: $fileSize bytes"
        
        # 1. 创建上传会话
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }
        
        $body = @{
            "@microsoft.graph.conflictBehavior" = "replace"  # 覆盖已存在的文件
            "name" = (Split-Path $FilePath -Leaf)
        } | ConvertTo-Json
        
        $session = Invoke-RestMethod -Uri $uploadUrl -Method Post -Headers $headers -Body $body
        $uploadUrl = $session.uploadUrl
        Write-Log "Upload session created: $uploadUrl"
        
        # 2. 分片上传
        $fileStream = [System.IO.File]::OpenRead($LocalFilePath)
        $bytesRemaining = $fileSize
        $offset = 0
        
        while ($bytesRemaining -gt 0) {
            $currentChunkSize = [Math]::Min($ChunkSize, $bytesRemaining)
            $buffer = New-Object byte[] $currentChunkSize
            $fileStream.Read($buffer, 0, $currentChunkSize) | Out-Null
            
            # 设置分片范围
            $endOffset = $offset + $currentChunkSize - 1
            $contentRange = "bytes $offset-$endOffset/$fileSize"
            
            $chunkHeaders = @{
                "Authorization" = "Bearer $Token"
                "Content-Length" = "$currentChunkSize"
                "Content-Range" = $contentRange
            }
            
            Write-Log "Uploading chunk: $contentRange"
            
            # 上传分片
            $response = Invoke-RestMethod -Uri $uploadUrl -Method Put -Headers $chunkHeaders -Body $buffer
            
            $offset += $currentChunkSize
            $bytesRemaining -= $currentChunkSize
            
            # 显示进度
            $percentComplete = [Math]::Round(($offset / $fileSize) * 100, 2)
            Write-Log "Progress: $percentComplete% ($offset / $fileSize bytes)"
        }
        
        $fileStream.Close()
        Write-Log "File uploaded successfully. File ID: $($response.id)"
        return $response
    }
    catch {
        Write-Log "Failed to upload large file: $_" "ERROR"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Log "Response body: $responseBody" "ERROR"
        }
        return $null
    }
}

function upload_file_to_sharepoint{
    param(
        [string]$Token,
        [string]$LocalFilePath,
        [string]$DestPath
    )
    
#https://graph.microsoft.com/v1.0/sites/riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93/drives
    $SiteID="riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93"
    $DriveID="b!CtY3L4EriEqd7iiLX8JZ8sifkBaHrZ9KlsIi3K1ICpPR4Y7ssFseQpJlF9TBR2Yi"
    Upload-LargeFileToSharePoint -SiteID $SiteID `
        -DriveID $DriveID `
        -FilePath "R&D/VideoEngine/$DestPath" `
        -LocalFilePath $LocalFilePath `
        -Token $Token
}
# 主程序
function Main {
    Setup-Logger
    
    Write-Log "SharePoint Downloader Started"
    
    # 加载配置文件
    $config = Load-Config -ConfigPath $ConfigFile
    
    # 获取 Azure 凭证
    $azureClientID = $null
    $azureClientSecret = $null
    $azureTenantID = $null
    $urlList = @()
    
    if ($config) {
        # 从配置文件获取凭证
        $azureClientID = $config.azure_client_id
        $azureClientSecret = $config.azure_client_secret
        $azureTenantID = $config.azure_tenant_id
        
        if ($config.sharepoint_url) {
            # 处理 sharepoint_url 数组
            if ($config.sharepoint_url -is [array]) {
                foreach ($item in $config.sharepoint_url) {
                    if ($item.url) {
                        $urlList += @{ name = $item.name; url = $item.url }
                    }
                }
            }
            elseif ($config.sharepoint_url -is [string]) {
                $urlList += @{ name = ""; url = $config.sharepoint_url }
            }
        }
        
        if ($SaveDir -eq ".") {
            if($config.download_folder -eq $null) {
                $SaveDir = $PSScriptRoot
            } else{
                $SaveDir = $config.download_folder 
            }

            
        }
    }
    
    # 如果没有从配置文件获取到凭证，尝试从命令行参数和环境变量
    if (-not $azureClientID -or -not $azureClientSecret -or -not $azureTenantID) {
        Write-Log "Loading credentials from environment variables..."
        LoadEnv
        
        if($azureClientID -eq $null){
            $azureClientID = $env:AZURE_CLIENT_ID
        }
        if( $azureClientSecret -eq $null){
            $azureClientSecret = $env:AZURE_CLIENT_SECRET
        }
        if( $azureTenantID -eq $null){
            $azureTenantID = $env:AZURE_TENANT_ID
        }
    }
    
    # 如果没有URL从配置文件中获取，使用命令行参数
    if ($urlList.Count -eq 0 -and $SharePointURL) {
        $urlList += @{ name = ""; url = $SharePointURL }
    }
    
    Write-Log "Save Directory: $SaveDir"
    
    # 验证凭证
    if (-not $azureClientID -or -not $azureClientSecret -or -not $azureTenantID) {
        Write-Log "Missing Azure credentials" "ERROR"
        return $false
    }
    
    # 验证保存目录
    if (-not (Test-Path -Path $SaveDir)) {
        Write-Log "Creating directory: $SaveDir"
        New-Item -ItemType Directory -Path $SaveDir | Out-Null
    }
    
    # 获取 Access Token
    Write-Log "Getting Access Token..."
    $tokenURL = "https://login.microsoftonline.com/$azureTenantID/oauth2/v2.0/token"
    $tokenBody = @{
        client_id = $azureClientID
        scope = "https://graph.microsoft.com/.default"
        client_secret = $azureClientSecret
        grant_type = "client_credentials"
    }
    
    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenURL -Method Post -Body $tokenBody
        $token = $tokenResponse.access_token
        Write-Log "Access token acquired successfully"
        Write-Log "Token (first 30 chars): $($token.Substring(0, 30))..."
    }
    catch {
        Write-Log "Failed to get access token: $_" "ERROR"
        return $false
    }
    
# Testupload
if($UploadLocalFile -and $UploadDestPath) {
    upload_file_to_sharepoint -Token $token -LocalFilePath $UploadLocalFile -DestPath $UploadDestPath
    return $true
}
    # 下载文件列表中的所有 URL
    $allSuccess = $true
    if ($urlList.Count -eq 0) {
        Write-Log "No SharePoint URLs to download" "ERROR"
        return $false
    }
    
    foreach ($item in $urlList) {
        $url = $item.url
        $name = $item.name
        if (-not $name) {
            $name = Split-Path -Leaf $url
        }
        
        Write-Log "Processing URL: $url"
        $result = Download-FromSharePoint -SharedURL $url -SaveDir $SaveDir -Token $token
        
        if ($result) {
            Write-Log "$name downloaded successfully"
        }
        else {
            Write-Log "$name download failed" "ERROR"
            $allSuccess = $false
        }
    }
    
    if ($allSuccess) {
        Write-Log "Program completed successfully at $(Get-Date)"
        return $true
    }
    else {
        Write-Log "Program completed with some failures at $(Get-Date)" "WARNING"
        return $false
    }
}

# 执行主程序
$success = Main
exit $(if ($success) { 0 } else { 1 })
