# =================== CONFIGURATION ===================
$RecurseSubDirs = $true  # Set to $false to skip subdirectories

# SFTP Credentials and Path
$SftpHost = "sftp.internal.company"
$SftpUsername = "sftp-service-user"
$SftpPassword = "sftp-password"
$SftpRemotePath = "/source/data"

# Azure SMB Credentials and Path
$SmbPath = "\\azurestorage.file.core.windows.net\myshare"
$SmbUsername = "DOMAIN\service-user"
$SmbPassword = "service-user-password"

# Path to WinSCP .NET assembly
$WinScpDllPath = "C:\Tools\WinSCP\WinSCPnet.dll"

# Local temp download directory
$LocalTempDir = "C:\Temp\SftpDownload"
# ====================================================

# Ensure local temp directory exists
if (!(Test-Path $LocalTempDir)) {
    New-Item -ItemType Directory -Path $LocalTempDir | Out-Null
}

# Load WinSCP .NET assembly
if (!(Test-Path $WinScpDllPath)) {
    Write-Error "WinSCP .NET assembly not found at $WinScpDllPath"
    exit 1
}
Add-Type -Path $WinScpDllPath

# Create session and transfer options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = $SftpHost
    UserName = $SftpUsername
    Password = $SftpPassword
    SshHostKeyPolicy = "GiveUpSecurityAndAcceptAny"  # Consider importing host key for security
}

$session = New-Object WinSCP.Session
try {
    $session.Open($sessionOptions)

    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

    if ($RecurseSubDirs) {
        $transferOptions.FileMask = "*.*"
    } else {
        $transferOptions.FileMask = "*.*|*/"  # Exclude subdirectories
    }

    Write-Host "Downloading from SFTP..."
    $transferResult = $session.GetFiles($SftpRemotePath, "$LocalTempDir\*", $false, $transferOptions)
    $transferResult.Check()  # Throw if any fail

    Write-Host "Connecting to Azure SMB..."

    # Map network drive (optional: skip if already mounted)
    $netResource = @{
        LocalName = ""
        RemoteName = $SmbPath
        ProviderName = "Microsoft Windows Network"
    }

    $securePassword = ConvertTo-SecureString $SmbPassword -AsPlainText -Force
    $credentials = New-Object System.Management.Automation.PSCredential($SmbUsername, $securePassword)

    # Copy to Azure SMB
    Write-Host "Copying files to SMB share..."
    $files = Get-ChildItem -Path $LocalTempDir -Recurse:$RecurseSubDirs
    foreach ($file in $files) {
        $destPath = Join-Path $SmbPath ($file.FullName.Substring($LocalTempDir.Length).TrimStart("\"))
        $destDir = Split-Path $destPath

        if (!(Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        Copy-Item -Path $file.FullName -Destination $destPath -Force
    }

    Write-Host "Transfer complete."

} catch {
    Write-Error "Error during transfer: $_"
} finally {
    $session.Dispose()
}
