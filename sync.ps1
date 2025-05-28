###########################################################
# PowerShell Script: SFTP Transfer with WinSCP .NET       #
# Requirements: WinSCP .NET assembly (WinSCPnet.dll)       #
# https://winscp.net/eng/docs/library                     #
###########################################################

# Load WinSCP .NET Assembly
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

# ====== CONFIGURATION VARIABLES ======
# --- SFTP Credentials and Connection Info ---
$SftpHostA = "sftp.remote-location-a.com"
$SftpUserA = "usernameA"
$SftpPassA = "passwordA"
$SshKeyFingerprintA = "ssh-rsa 2048 xx:xx:xx:xx..." # Move to secure store for production

$SftpHostB = "sftp.remote-location-b.com"
$SftpUserB = "usernameB"
$SftpPassB = "passwordB"
$SshKeyFingerprintB = "ssh-rsa 2048 yy:yy:yy:yy..." # Move to secure store for production

# --- Pathing Info ---
$NetworkPathA     = "\\NetworkShareA\Uploads"
$RemotePathA      = "/uploads"

$RemotePathB      = "/downloads"
$NetworkPathB     = "\\NetworkShareB\Downloads"

$ZipFilePath      = "$env:TEMP\upload.zip"
$LogDirectory     = "\\NetworkShareLogs\SftpLogs"
$LogFilePath      = Join-Path $LogDirectory ("sftp_transfer_log_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")

# --- SMTP Info for Summary Email ---
$SmtpServer = "smtp.example.com"
$SmtpPort = 587
$SmtpUser = "smtpuser"
$SmtpPass = "smtppass"
$EmailTo = "admin@example.com"
$EmailFrom = "automation@sftp.com"

# ====== ZIP SECTION (BEFORE UPLOAD TO REMOTE A) ======
# Create or clear log file
"Transfer Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" | Out-File -FilePath $LogFilePath -Encoding UTF8

# Log each file being added to zip
Get-ChildItem -Path $NetworkPathA -File | ForEach-Object {
    "Added to zip: $($_.FullName) at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# Compress contents of Network Location A into a zip
Compress-Archive -Path "$NetworkPathA\*" -DestinationPath $ZipFilePath -Force

# ====== UPLOAD TO REMOTE LOCATION A ======
# Create SFTP Session Options for Remote A
$sessionOptionsA = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = $SftpHostA
    UserName = $SftpUserA
    Password = $SftpPassA
    SshHostKeyFingerprint = $SshKeyFingerprintA
}

$sessionA = New-Object WinSCP.Session
$uploadSuccess = $false
try {
    $sessionA.Open($sessionOptionsA)

    # Upload the zip file to Remote A
    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

    $transferResult = $sessionA.PutFiles($ZipFilePath, $RemotePathA + "/", $False, $transferOptions)
    $transferResult.Check()
    $uploadCount = $transferResult.Transfers.Count
    $uploadSuccess = $true
}
catch {
    $uploadError = $_.Exception.Message
}
finally {
    $sessionA.Dispose()
}

# Remove source files after upload to prevent reprocessing
if ($uploadSuccess) {
    Get-ChildItem -Path $NetworkPathA -File | Remove-Item -Force
    "Source files deleted after upload." | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# ====== DOWNLOAD FROM REMOTE LOCATION B ======
# Create SFTP Session Options for Remote B
$sessionOptionsB = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = $SftpHostB
    UserName = $SftpUserB
    Password = $SftpPassB
    SshHostKeyFingerprint = $SshKeyFingerprintB
}

$sessionB = New-Object WinSCP.Session
$downloadSuccess = $false
try {
    $sessionB.Open($sessionOptionsB)

    # List files in remote directory to find latest zip
    $directoryInfo = $sessionB.ListDirectory($RemotePathB)
    $zipFile = $directoryInfo.Files | Where-Object { !$_.IsDirectory -and $_.Name -like "*.zip" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($zipFile -ne $null) {
        $remoteZipFile = $RemotePathB + "/" + $zipFile.Name
        $localZipFile = "$env:TEMP\download.zip"

        $transferResult = $sessionB.GetFiles($remoteZipFile, $localZipFile, $False)
        $transferResult.Check()
        $downloadCount = $transferResult.Transfers.Count
        $downloadSuccess = $true

        # Log the file that was downloaded
        "Downloaded file: $remoteZipFile at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    } else {
        throw "No zip file found in remote location B."
    }
}
catch {
    $downloadError = $_.Exception.Message
}
finally {
    $sessionB.Dispose()
}

# ====== UNZIP SECTION (AFTER DOWNLOAD FROM REMOTE B) ======
# Extract downloaded zip to Network Location B
if (Test-Path $localZipFile) {
    Expand-Archive -Path $localZipFile -DestinationPath $NetworkPathB -Force
}

# ====== EMAIL SUMMARY ======
$summaryBody = "SFTP Transfer Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
$summaryBody += "Upload to $SftpHostA: " + ($uploadSuccess ? "Success ($uploadCount files)" : "Failed - $uploadError") + "`n"
$summaryBody += "Download from $SftpHostB: " + ($downloadSuccess ? "Success ($downloadCount files)" : "Failed - $downloadError") + "`n"

$mailParams = @{
    SmtpServer = $SmtpServer
    Port = $SmtpPort
    From = $EmailFrom
    To = $EmailTo
    Subject = "SFTP Transfer Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Body = $summaryBody
    Attachments = $LogFilePath
    Credential = New-Object System.Management.Automation.PSCredential($SmtpUser, (ConvertTo-SecureString $SmtpPass -AsPlainText -Force))
    UseSsl = $true
}

Send-MailMessage @mailParams

Write-Host "SFTP Transfer and File Processing Completed. Summary email sent."
