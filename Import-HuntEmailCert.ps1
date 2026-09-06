<#
.SYNOPSIS
    One-time import of the toolkit PFX into Current User\My.

.DESCRIPTION
    Run this once per machine (or after cert rotation) to import the .pfx so that
    Connect-HuntEmailGraph.ps1 can connect without prompting for password each time.

    PFX location: C:\certs\huntEmail.pfx (from cert\createCert.ps1), or pass -PfxPath.

.PARAMETER PfxPath
    Path to the .pfx file. Default: C:\certs\huntEmail.pfx

.PARAMETER Password
    SecureString password for the PFX (for example, from your secret manager).

.EXAMPLE
    $pwd = Read-Host "PFX password" -AsSecureString
    .\Import-HuntEmailCert.ps1 -Password $pwd

.EXAMPLE
    .\Import-HuntEmailCert.ps1 -PfxPath "C:\certs\huntEmail.pfx" -Password (Read-Host "Password" -AsSecureString)
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$PfxPath = "C:\certs\huntEmail.pfx",

    [Parameter(Mandatory)]
    [SecureString]$Password
)

if (-not (Test-Path -LiteralPath $PfxPath)) {
    Write-Host "[ERROR] PFX not found: $PfxPath" -ForegroundColor Red
    Write-Host "  Create it with cert\createCert.ps1, or copy from your secure location." -ForegroundColor Gray
    exit 1
}

Write-Host "[CERT] Importing to Cert:\CurrentUser\My..." -ForegroundColor Cyan
try {
    $imported = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password -ErrorAction Stop
    Write-Host "[CERT] Imported. Thumbprint: $($imported.Thumbprint)" -ForegroundColor Green
    Write-Host "  Next: . .\Connect-HuntEmailGraph.ps1" -ForegroundColor Gray
}
catch {
    Write-Host "[ERROR] Import failed: $_" -ForegroundColor Red
    exit 1
}
