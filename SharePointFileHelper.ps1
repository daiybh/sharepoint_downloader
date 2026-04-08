

# =====================
# SharePoint File Helper
# Support SharePoint file upload/download, config file and parameterization
# =====================

param(
    [Parameter()][string]$SharePointURL = "",           # Single download URL
    [Parameter()][string]$SaveFileName = "",            # Downloaded file name
    [Parameter()][string]$UploadLocalFile = "",         # Local file path to upload
    [Parameter()][string]$UploadDestPath = "",          # Destination path in SharePoint
    [Parameter()][string]$SiteID = "riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93",                  # SharePoint SiteID (optional, parameter priority)
    [Parameter()][string]$DriveID = "b!CtY3L4EriEqd7iiLX8JZ8sifkBaHrZ9KlsIi3K1ICpPR4Y7ssFseQpJlF9TBR2Yi",                 # SharePoint DriveID (optional, parameter priority)
    [Parameter()][bool]$EnableLogging = $true,           # Enable logging
    [Parameter()][string]$SaveDir = ".",               # Download save directory
    [Parameter()][string]$LogFile = "sharepoint_downloader.log", # Log file name
    [Parameter()][string]$ConfigFile = "config.json"    # Config file path
)

# ========== Logging ==========
$script:LogFile = $LogFile
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "ERROR", "WARNING")][string]$Level = "INFO"
    )
    # Write log message to both console and log file if logging is enabled
    if ($EnableLogging) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp $Level $Message"
        Write-Host $logMessage
        Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8
    }
}
function Setup-Logger {
    # Create log file if logging is enabled and file does not exist
    if ($EnableLogging -and -not (Test-Path -Path $script:LogFile)) {
        New-Item -ItemType File -Path $script:LogFile | Out-Null
    }
}

# Parse SharePoint URL
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

# Get SiteID
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

# Get DriveID
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

# Get file information
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

# Download file
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


# Get SiteID, DriveID, and FileInfo from SharePoint URL
function Get-SiteID-DriveID-FromURL {
    param(
        [string]$SharedURL,
        [string]$Token
    )
    # Parse URL
    $urlParts = Split-SharePointURL -SharedURL $SharedURL
    if (-not $urlParts) {
        Write-Log "URL parsing failed." "ERROR"
        return $null
    }
    $domain = $urlParts.Domain
    $sitePath = $urlParts.SitePath
    $filePath = $urlParts.FilePath
    Write-Log "Domain: $domain, SitePath: $sitePath, FilePath: $filePath"
    # Get SiteID
    $siteID = Get-SiteID -Domain $domain -SitePath $sitePath -Token $Token
    if (-not $siteID) {
        Write-Log "Unable to get siteID." "ERROR"
        return $null
    }
    Write-Log "SiteID: $siteID"
    # Get DriveID
    $driveID = Get-DriveID -SiteID $siteID -Token $Token
    if (-not $driveID) {
        Write-Log "Unable to get driveID." "ERROR"
        return $null
    }
    Write-Log "DriveID: $driveID"
    # Get file information
    $fileInfo = Get-FileInfo -SiteID $siteID -DriveID $driveID -FilePath $filePath -Token $Token
    if (-not $fileInfo ) {
        Write-Log "Unable to get file info." "ERROR"
        return $null
    }
    Write-Log "FileInfo: $($fileInfo | ConvertTo-Json -Depth 5)"
    return @{ SiteID = $siteID; DriveID = $driveID; FileInfo = $fileInfo; FilePath = $filePath }
}
# Download from SharePoint
function Download-FromSharePoint {
    param(
        [string]$SharedURL,
        [string]$SaveDir,
        [string]$Token
    )
    # Get SiteID, DriveID and file info using new return object
    $result = Get-SiteID-DriveID-FromURL -SharedURL $SharedURL -Token $Token
    if (-not $result) {
        Write-Log "Failed to get siteID, driveID and file info, aborting download." "ERROR"
        return $false
    }
    $siteID = $result.SiteID
    $driveID = $result.DriveID
    $fileInfo = $result.FileInfo
    $filePath = $result.FilePath
    Write-Log "siteID: $siteID, driveID: $driveID"
    if (-not $fileInfo -or -not $fileInfo.'@microsoft.graph.downloadUrl') {
        Write-Log "Unable to get download link, aborting download." "ERROR"
        return $false
    }
    $downloadURL = $fileInfo.'@microsoft.graph.downloadUrl'
    Write-Log "@microsoft.graph.downloadUrl $downloadURL"
    # Prepare local path
    $fileName = Split-Path -Leaf $filePath
    $localPath = Join-Path $SaveDir "$fileName"
    if ($SaveFileName) {
        $localPath = Join-Path $SaveDir $SaveFileName
    }
    # Download file
    return Download-File -DownloadURL $downloadURL -LocalPath $localPath
}

function LoadEnv {
    param (
        [string]$Path = ".env"
    )
    # Read file and iterate through each line
    Get-Content $Path | ForEach-Object {
        # Ignore empty lines and comment lines
        if ($_ -and $_ -notmatch '^\s*#') {
            # Split key-value pairs by the first equals sign
            $key, $value = $_ -split '=', 2
            # Set environment variable
            Set-Item -Path "env:$key" -Value $value
        }
    }    
}

# Read configuration file
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
        [string]$FilePath,        # Target path in SharePoint
        [string]$LocalFilePath,   # Local file path
        [string]$Token,
        [int]$ChunkSize = 5 * 1024 * 1024  # Default 5MB per chunk
    )

    $uploadUrl = "https://graph.microsoft.com/v1.0/sites/${SiteID}/drives/${DriveID}/root:/${FilePath}:/createUploadSession"
    Write-Log "####################"
    Write-Log "upload_large_file_to_sharepoint"
    Write-Log "Target URL: $uploadUrl"
    Write-Log "Local file: $LocalFilePath"

    try {
        # Check local file
        if (-not (Test-Path $LocalFilePath)) {
            Write-Log "Local file not found: $LocalFilePath" "ERROR"
            return $null
        }

        $fileInfo = Get-Item $LocalFilePath
        $fileSize = $fileInfo.Length
        Write-Log "File size: $fileSize bytes"

        # 1. Create upload session
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }

        $body = @{
            "@microsoft.graph.conflictBehavior" = "replace"  # option: replace, rename, fail
            "name" = (Split-Path $FilePath -Leaf)
        } | ConvertTo-Json

        $session = Invoke-RestMethod -Uri $uploadUrl -Method Post -Headers $headers -Body $body
        $uploadUrl = $session.uploadUrl
        Write-Log "Upload session created: $uploadUrl"

        # 2. Chunked upload
        $fileStream = [System.IO.File]::OpenRead($LocalFilePath)
        $bytesRemaining = $fileSize
        $offset = 0

        while ($bytesRemaining -gt 0) {
            $currentChunkSize = [Math]::Min($ChunkSize, $bytesRemaining)
            $buffer = New-Object byte[] $currentChunkSize
            $fileStream.Read($buffer, 0, $currentChunkSize) | Out-Null

            # Set chunk range
            $endOffset = $offset + $currentChunkSize - 1
            $contentRange = "bytes $offset-$endOffset/$fileSize"

            $chunkHeaders = @{
                "Authorization" = "Bearer $Token"
                "Content-Length" = "$currentChunkSize"
                "Content-Range" = $contentRange
            }

            Write-Log "Uploading chunk: $contentRange"

            # Upload chunk
            $response = Invoke-RestMethod -Uri $uploadUrl -Method Put -Headers $chunkHeaders -Body $buffer

            $offset += $currentChunkSize
            $bytesRemaining -= $currentChunkSize

            # Show progress
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


# Upload file to SharePoint (SiteID/DriveID passed as parameter)
function Upload-FileToSharePoint {
    param(
        [string]$Token,
        [string]$LocalFilePath,
        [string]$DestPath,
        [string]$SiteID,
        [string]$DriveID
    )
    if (-not $SiteID -or -not $DriveID) {
        Write-Log "SiteID/DriveID not specified, cannot upload." "ERROR"
        return $null
    }
    return Upload-LargeFileToSharePoint -SiteID $SiteID `
        -DriveID $DriveID `
        -FilePath $DestPath `
        -LocalFilePath $LocalFilePath `
        -Token $Token
}


# ========== Main Logic ==========
# Entry point for the script
function Main {
    Setup-Logger
    Write-Log "***************"
    Write-Log "Program started at $(Get-Date)"
    try {
        # Load config and initialize settings
        $config = Load-Config -ConfigPath $ConfigFile
        $settings = Initialize-Settings -Config $config
        # Get Azure AD access token
        $token = Get-AccessToken -Settings $settings
        if (-not $token) {
            Write-Log "Failed to get access token" "ERROR"
            return $false
        }
        # Upload has higher priority
        if ($UploadLocalFile -and $UploadDestPath) {
            $siteIdToUse = $SiteID; $driveIdToUse = $DriveID
            if (-not $siteIdToUse) { $siteIdToUse = $settings.siteID }
            if (-not $driveIdToUse) { $driveIdToUse = $settings.driveID }
            return Handle-Upload -Token $token -LocalFile $UploadLocalFile -DestPath $UploadDestPath -SiteID $siteIdToUse -DriveID $driveIdToUse
        } else {
            return Handle-Downloads -Token $token -Urls $settings.urlList -SaveDir $settings.saveDir
        }
    } catch {
        Write-Log "An unexpected error occurred: $_" "ERROR"
        return $false
    } finally {
        Write-Log "Program completed at $(Get-Date)"
    }
}


# Initialize config and parameters
function Initialize-Settings {
    param([object]$Config)
    $settings = @{
        azureClientID = $null
        azureClientSecret = $null
        azureTenantID = $null
        urlList = @()
        saveDir = $SaveDir
        siteID = $null
        driveID = $null
    }
    if ($Config) {
        $settings.azureClientID = $Config.azure_client_id
        $settings.azureClientSecret = $Config.azure_client_secret
        $settings.azureTenantID = $Config.azure_tenant_id
        if ($Config.site_id) { $settings.siteID = $Config.site_id }
        if ($Config.drive_id) { $settings.driveID = $Config.drive_id }
        # Parse SharePoint URL list from config
        if ($Config.sharepoint_url) {
            if ($Config.sharepoint_url -is [array]) {
                foreach ($item in $Config.sharepoint_url) {
                    if ($item.url) { $settings.urlList += @{ name = $item.name; url = $item.url } }
                }
            } elseif ($Config.sharepoint_url -is [string]) {
                $settings.urlList += @{ name = ""; url = $Config.sharepoint_url }
            }
        }
        # Set download folder from config if not specified
        if ($settings.saveDir -eq "." -and $Config.download_folder) {
            $settings.saveDir = $Config.download_folder
        }
    }
    # Fallback to environment variables
    if (-not $settings.azureClientID -or -not $settings.azureClientSecret -or -not $settings.azureTenantID) {
        Write-Log "Loading credentials from environment variables..."
        LoadEnv
        if (-not $settings.azureClientID) { $settings.azureClientID = $env:AZURE_CLIENT_ID }
        if (-not $settings.azureClientSecret) { $settings.azureClientSecret = $env:AZURE_CLIENT_SECRET }
        if (-not $settings.azureTenantID) { $settings.azureTenantID = $env:AZURE_TENANT_ID }
    }
    # Fallback to command line parameter
    if ($settings.urlList.Count -eq 0 -and $SharePointURL) {
        $settings.urlList += @{ name = ""; url = $SharePointURL }
    }
    if (-not $settings.siteID -and $SiteID) { $settings.siteID = $SiteID }
    if (-not $settings.driveID -and $DriveID) { $settings.driveID = $DriveID }
    if (-not $settings.azureClientID -or -not $settings.azureClientSecret -or -not $settings.azureTenantID) {
        throw "Missing Azure credentials"
    }
    # Ensure download directory exists
    if (-not (Test-Path -Path $settings.saveDir)) {
        Write-Log "Creating directory: $settings.saveDir"
        New-Item -ItemType Directory -Path $settings.saveDir | Out-Null
    }
    Write-Log "Save Directory: $settings.saveDir"
    return $settings
}

# Get access token
function Get-AccessToken {
    param(
        [hashtable]$Settings
    )

    Write-Log "Getting Access Token..."
    $tokenURL = "https://login.microsoftonline.com/$($settings.azureTenantID)/oauth2/v2.0/token"
    $tokenBody = @{
        client_id = $settings.azureClientID
        scope = "https://graph.microsoft.com/.default"
        client_secret = $settings.azureClientSecret
        grant_type = "client_credentials"
    }

    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenURL -Method Post -Body $tokenBody
        $token = $tokenResponse.access_token
        Write-Log "Access token acquired successfully"
        Write-Log "Token (first 30 chars): $($token.Substring(0, 30))..."
        return $token
    }
    catch {
        Write-Log "Failed to get access token: $_" "ERROR"
        return $null
    }
}


# Upload operation wrapper
function Handle-Upload {
    param(
        [string]$Token,
        [string]$LocalFile,
        [string]$DestPath,
        [string]$SiteID,
        [string]$DriveID
    )
    Write-Log "Starting upload operation"
    $response = Upload-FileToSharePoint -Token $Token -LocalFilePath $LocalFile -DestPath $DestPath -SiteID $SiteID -DriveID $DriveID
    if ($response) {
        Write-Log "File uploaded successfully. File ID: $($response.id)"
        return $true
    } else {
        Write-Log "File upload failed." "ERROR"
        return $false
    }
}

# Handle download operation
function Handle-Downloads {
    param(
        [string]$Token,
        [array]$Urls,
        [string]$SaveDir
    )

    if ($Urls.Count -eq 0) {
        Write-Log "No SharePoint URLs to download" "ERROR"
        return $false
    }

    $allSuccess = $true

    foreach ($item in $Urls) {
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
        Write-Log "All files downloaded successfully"
        return $true
    }
    else {
        Write-Log "Some files failed to download" "WARNING"
        return $false
    }
}

# Execute main program
$success = Main
exit $(if ($success) { 0 } else { 1 })
