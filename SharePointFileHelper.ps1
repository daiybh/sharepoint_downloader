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

# Global variables
$script:LogFile = $LogFile
# Setup logger
function Setup-Logger {
    if ($EnableLogging) {     
        if (-not (Test-Path -Path $script:LogFile)) {
            New-Item -ItemType File -Path $script:LogFile | Out-Null
        }
    }
}

# Log output to file and console
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

# Download from SharePoint
function Download-FromSharePoint {
    param(
        [string]$SharedURL,
        [string]$SaveDir,
        [string]$Token
    )

    # Parse URL
    $urlParts = Split-SharePointURL -SharedURL $SharedURL
    if (-not $urlParts) {
        Write-Log "URL parsing failed, aborting download." "ERROR"
        return $false
    }

    $domain = $urlParts.Domain
    $sitePath = $urlParts.SitePath
    $filePath = $urlParts.FilePath
    Write-Log "Domain: $domain, SitePath: $sitePath, FilePath: $filePath"

    # Get SiteID
    $siteID = Get-SiteID -Domain $domain -SitePath $sitePath -Token $Token
    if (-not $siteID) {
        Write-Log "Unable to get siteID, aborting download." "ERROR"
        return $false
    }
    Write-Log "SiteID: $siteID"

    # Get DriveID
    $driveID = Get-DriveID -SiteID $siteID -Token $Token
    if (-not $driveID) {
        Write-Log "Unable to get driveID, aborting download." "ERROR"
        return $false
    }
    Write-Log "DriveID: $driveID"

    # Get file information
    $fileInfo = Get-FileInfo -SiteID $siteID -DriveID $driveID -FilePath $filePath -Token $Token
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
            "@microsoft.graph.conflictBehavior" = "replace"  # Overwrite existing file
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

function upload_file_to_sharepoint{
    param(
        [string]$Token,
        [string]$LocalFilePath,
        [string]$DestPath
    )
    $SiteID="riedelcommunications.sharepoint.com,2f37d60a-2b81-4a88-9dee-288b5fc259f2,16909fc8-ad87-4a9f-96c2-22dcad480a93"
    $DriveID="b!CtY3L4EriEqd7iiLX8JZ8sifkBaHrZ9KlsIi3K1ICpPR4Y7ssFseQpJlF9TBR2Yi"
    return Upload-LargeFileToSharePoint -SiteID $SiteID `
        -DriveID $DriveID `
        -FilePath "R&D/VideoEngine/$DestPath" `
        -LocalFilePath $LocalFilePath `
        -Token $Token
}

# Main program
function Main {
    Setup-Logger
    Write-Log "***************"
    Write-Log "Program started at $(Get-Date)"

    try {
        # Load configuration and settings
        $config = Load-Config -ConfigPath $ConfigFile
        $settings = Initialize-Settings -Config $config

        # Get access token
        $token = Get-AccessToken -Settings $settings
        if (-not $token) {
            Write-Log "Failed to get access token" "ERROR"
            return $false
        }

        # Execute appropriate functionality based on operation mode
        if ($UploadLocalFile -and $UploadDestPath) {
            return Handle-Upload -Token $token -LocalFile $UploadLocalFile -DestPath $UploadDestPath
        } else {
            return Handle-Downloads -Token $token -Urls $settings.urlList -SaveDir $settings.saveDir
        }
    }
    catch {
        Write-Log "An unexpected error occurred: $_" "ERROR"
        return $false
    }
    finally {
        Write-Log "Program completed at $(Get-Date)"
    }
}

# Initialize settings
function Initialize-Settings {
    param(
        [object]$Config
    )

    $settings = @{
        azureClientID = $null
        azureClientSecret = $null
        azureTenantID = $null
        urlList = @()
        saveDir = $SaveDir
    }

    # Get settings from config file
    if ($Config) {
        $settings.azureClientID = $Config.azure_client_id
        $settings.azureClientSecret = $Config.azure_client_secret
        $settings.azureTenantID = $Config.azure_tenant_id

        # Process SharePoint URL list
        if ($Config.sharepoint_url) {
            if ($Config.sharepoint_url -is [array]) {
                foreach ($item in $Config.sharepoint_url) {
                    if ($item.url) {
                        $settings.urlList += @{ name = $item.name; url = $item.url }
                    }
                }
            }
            elseif ($Config.sharepoint_url -is [string]) {
                $settings.urlList += @{ name = ""; url = $Config.sharepoint_url }
            }
        }

        # Set save directory
        if ($settings.saveDir -eq "." -and $Config.download_folder) {
            $settings.saveDir = $Config.download_folder
        }
    }

    # If credentials were not obtained from config file, try from environment variables
    if (-not $settings.azureClientID -or -not $settings.azureClientSecret -or -not $settings.azureTenantID) {
        Write-Log "Loading credentials from environment variables..."
        LoadEnv

        if (-not $settings.azureClientID) {
            $settings.azureClientID = $env:AZURE_CLIENT_ID
        }
        if (-not $settings.azureClientSecret) {
            $settings.azureClientSecret = $env:AZURE_CLIENT_SECRET
        }
        if (-not $settings.azureTenantID) {
            $settings.azureTenantID = $env:AZURE_TENANT_ID
        }
    }

    # If no URLs from config file, use command line parameter
    if ($settings.urlList.Count -eq 0 -and $SharePointURL) {
        $settings.urlList += @{ name = ""; url = $SharePointURL }
    }

    # Validate settings
    if (-not $settings.azureClientID -or -not $settings.azureClientSecret -or -not $settings.azureTenantID) {
        throw "Missing Azure credentials"
    }

    # Ensure save directory exists
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

# Handle upload operation
function Handle-Upload {
    param(
        [string]$Token,
        [string]$LocalFile,
        [string]$DestPath
    )

    Write-Log "Starting upload operation"
    $response = upload_file_to_sharepoint -Token $token -LocalFilePath $LocalFile -DestPath $DestPath

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
