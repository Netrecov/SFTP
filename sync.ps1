# Load WinSCP .NET assembly
Add-Type -Path "C:\Path\To\WinSCPnet.dll"

# ===== Configuration Variables =====
# SFTP connection
$SftpHost = "sftp.remote-site.com"
$SftpUser = "sftpuser"
$SftpPass = "sftppass"
$SshKeyFingerprint = "ssh-rsa 2048 xx:xx:xx:xx:xx..."

# Remote paths
$RemotePathA = "/user_upload"
$RemotePathB = "/user_download"

# Network share paths
$NetworkLocationA = "\\network\share\upload"
$NetworkLocationB = "\\network\share\download"

# Logging
$logTime = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFilePath = "\\network\share\logs\SftpTransfer_$logTime.log"

# SMTP config
$SmtpServer = "smtp.example.com"
$SmtpPort = 587
$SmtpUser = "smtpuser"
$SmtpPass = "smtppass"
$EmailTo = "admin@example.com"
$EmailFrom = "automation@sftp.com"

# ===== Logging Utility =====
function Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp `t $Message" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# ===== Function 1: Upload-ZipToSftp =====
function Upload-ZipToSftp {
    $zipFile = "$env:TEMP\upload_$logTime.zip"
    Log "Zipping files from $NetworkLocationA..."
    Compress-Archive -Path "$NetworkLocationA\*" -DestinationPath $zipFile -Force

    # Track filenames
    Get-ChildItem -Path $NetworkLocationA -File | ForEach-Object {
        Log "Prepared for upload: $($_.Name)"
    }

    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SftpHost
        UserName = $SftpUser
        Password = $SftpPass
        SshHostKeyFingerprint = $SshKeyFingerprint
    }

    $session = New-Object WinSCP.Session
    try {
        $session.Open($sessionOptions)
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        $result = $session.PutFiles($zipFile, "$RemotePathA/", $False, $transferOptions)
        $result.Check()
        Log "Uploaded $($result.Transfers.Count) file(s) to $RemotePathA"

        # Delete local source files after upload
        Get-ChildItem -Path $NetworkLocationA -File | Remove-Item -Force
        Log "Source files at $NetworkLocationA deleted."
        return @{ Success = $true; Count = $result.Transfers.Count }
    }
    catch {
        Log "Upload failed: $($_.Exception.Message)"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        $session.Dispose()
    }
}

# ===== Function 2: Download-ZipFromSftp =====
function Download-ZipFromSftp {
    $zipFile = "$env:TEMP\download_$logTime.zip"

    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SftpHost
        UserName = $SftpUser
        Password = $SftpPass
        SshHostKeyFingerprint = $SshKeyFingerprint
    }

    $session = New-Object WinSCP.Session
    try {
        $session.Open($sessionOptions)
        $remoteFiles = $session.ListDirectory($RemotePathB).Files | Where-Object { !$_.IsDirectory -and $_.Name -like "*.zip" }
        $latestFile = $remoteFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latestFile) {
            throw "No zip file found at $RemotePathB"
        }

        $remotePath = "$RemotePathB/$($latestFile.Name)"
        $result = $session.GetFiles($remotePath, $zipFile, $False)
        $result.Check()
        Log "Downloaded $($result.Transfers.Count) file(s) from $RemotePathB"

        Expand-Archive -Path $zipFile -DestinationPath $NetworkLocationB -Force
        Log "Extracted contents to $NetworkLocationB"

        # Delete the zip file from remote after download
        $session.RemoveFiles($remotePath)
        Log "Deleted $($latestFile.Name) from $RemotePathB"

        return @{ Success = $true; Count = $result.Transfers.Count }
    }
    catch {
        Log "Download failed: $($_.Exception.Message)"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        $session.Dispose()
    }
}

# ===== MAIN EXECUTION =====
$uploadResult = Upload-ZipToSftp
$downloadResult = Download-ZipFromSftp

# ===== Summary Email =====
$summary = @()
$summary += "SFTP Transfer Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
$summary += "Upload to $RemotePathA: " + ($uploadResult.Success ? "Success ($($uploadResult.Count) files)" : "Failed - $($uploadResult.Error)") + "`n"
$summary += "Download from $RemotePathB: " + ($downloadResult.Success ? "Success ($($downloadResult.Count) files)" : "Failed - $($downloadResult.Error)") + "`n"

$mailParams = @{
    SmtpServer  = $SmtpServer
    Port        = $SmtpPort
    From        = $EmailFrom
    To          = $EmailTo
    Subject     = "SFTP Transfer Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Body        = ($summary -join "`n")
    Attachments = $LogFilePath
    Credential  = New-Object System.Management.Automation.PSCredential($SmtpUser, (ConvertTo-SecureString $SmtpPass -AsPlainText -Force))
    UseSsl      = $true
}
Send-MailMessage @mailParams
Write-Host "SFTP process complete. Summary email sent."
