# UploadToRemote5.ps1
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

# === Config ===
$SftpHost = "sftp.example.com"
$SftpUser = "your_sftp_user"
$SftpPass = "your_sftp_pass"
$SshKeyFingerprint = "ssh-rsa 2048 xx:xx:xx:xx:xx"

$NetworkLocation5 = "\\path\to\network_location_5"
$RemoteLocation5 = "/remote_location_5"

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

    $result = $session.PutFiles("$NetworkLocation5\*", "$RemoteLocation5/", $False, $transferOptions)
    $result.Check()
    Write-Host "Uploaded $($result.Transfers.Count) file(s) to $RemoteLocation5"
}
catch {
    Write-Error "Upload failed: $($_.Exception.Message)"
}
finally {
    $session.Dispose()
}
