# DownloadFromRemote6.ps1
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

# === Config ===
$SftpHost = "sftp.example.com"
$SftpUser = "your_sftp_user"
$SftpPass = "your_sftp_pass"
$SshKeyFingerprint = "ssh-rsa 2048 xx:xx:xx:xx:xx"

$RemoteLocation6 = "/remote_location_6"
$NetworkLocation6 = "\\path\to\network_location_6"

# === Session Setup ===
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

    $result = $session.GetFiles("$RemoteLocation6/*", "$NetworkLocation6\", $False, $transferOptions)
    $result.Check()
    Write-Host "Downloaded $($result.Transfers.Count) file(s) to $NetworkLocation6"
}
catch {
    Write-Error "Download failed: $($_.Exception.Message)"
}
finally {
    $session.Dispose()
}
