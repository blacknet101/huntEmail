<#
.SYNOPSIS
    Connect to Microsoft Graph for the email investigation toolkit using certificate auth.

.DESCRIPTION
    Uses a Microsoft Entra app registration and certificate thumbprint supplied
    via parameters or environment variables. Run this once before using
    huntEmail, timelineEmail, or purgeEmail.

    If the cert is not yet in your Current User store, use -PfxPath and -Password
    to import it first (e.g. from C:\certs\huntEmail.pfx or project cert folder).

    Required values can be provided with parameters or these environment variables:
    HUNTEMAIL_TENANT_ID, HUNTEMAIL_CLIENT_ID, HUNTEMAIL_CERT_THUMBPRINT.

.PARAMETER PfxPath
    Optional. Full path to the .pfx file. If provided and the thumbprint cert is not
    in Cert:\CurrentUser\My, the script will import the PFX then connect.

.PARAMETER Password
    Optional. SecureString password for the PFX. Only used when -PfxPath is provided
    and the cert needs to be imported.

.EXAMPLE
    .\Connect-HuntEmailGraph.ps1
    # Cert already in store; connects using thumbprint from api_info.

.EXAMPLE
    $pwd = Read-Host "PFX password" -AsSecureString
    .\Connect-HuntEmailGraph.ps1 -TenantId "<tenant-id>" -ClientId "<client-id>" -Thumbprint "<thumbprint>" -PfxPath "C:\certs\huntEmail.pfx" -Password $pwd
    # Imports PFX if needed, then connects.

.NOTES
    See documentation\api_info.md for the configuration template.
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$TenantId = $env:HUNTEMAIL_TENANT_ID,

    [Parameter()]
    [string]$ClientId = $env:HUNTEMAIL_CLIENT_ID,

    [Parameter()]
    [string]$Thumbprint = $env:HUNTEMAIL_CERT_THUMBPRINT,

    [Parameter()]
    [string]$PfxPath,

    [Parameter()]
    [SecureString]$Password
)

if (-not $TenantId -or -not $ClientId -or -not $Thumbprint) {
    Write-Host "[ERROR] Missing Graph connection settings." -ForegroundColor Red
    Write-Host "  Provide -TenantId, -ClientId, and -Thumbprint" -ForegroundColor Yellow
    Write-Host "  or set HUNTEMAIL_TENANT_ID, HUNTEMAIL_CLIENT_ID, and HUNTEMAIL_CERT_THUMBPRINT." -ForegroundColor Yellow
    Write-Host "  See documentation\api_info.md for the configuration template." -ForegroundColor Gray
    exit 1
}

$certInStore = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }

if (-not $certInStore -and $PfxPath -and $Password) {
    if (-not (Test-Path -LiteralPath $PfxPath)) {
        Write-Host "[ERROR] PFX not found: $PfxPath" -ForegroundColor Red
        exit 1
    }
    Write-Host "[CERT] Importing PFX to Current User\My..." -ForegroundColor Cyan
    try {
        Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password | Out-Null
        Write-Host "[CERT] Imported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to import PFX: $_" -ForegroundColor Red
        exit 1
    }
}
elseif (-not $certInStore -and $PfxPath) {
    Write-Host "[ERROR] Cert with thumbprint $Thumbprint not in store. Provide -Password to import from: $PfxPath" -ForegroundColor Red
    exit 1
}
elseif (-not $certInStore) {
    $defaultPfx = "C:\certs\huntEmail.pfx"
    $projectPfx = Join-Path $PSScriptRoot "cert\huntEmail.pfx"
    Write-Host "[ERROR] Cert with thumbprint $Thumbprint not in Cert:\CurrentUser\My." -ForegroundColor Red
    Write-Host "  Import first, e.g.:" -ForegroundColor Yellow
    Write-Host "    `$pwd = Read-Host 'PFX password' -AsSecureString" -ForegroundColor Gray
    Write-Host "    Import-PfxCertificate -FilePath '$defaultPfx' -CertStoreLocation Cert:\CurrentUser\My -Password `$pwd" -ForegroundColor Gray
    Write-Host "  Then run this script again, or use: .\Connect-HuntEmailGraph.ps1 -PfxPath '$defaultPfx' -Password `$pwd" -ForegroundColor Gray
    exit 1
}

Write-Host "[GRAPH] Connecting to Microsoft Graph (certificate auth)..." -ForegroundColor Cyan
try {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Host "[GRAPH] Connected. ClientId: $($ctx.ClientId), AuthType: $($ctx.AuthType)" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Connect-MgGraph failed: $_" -ForegroundColor Red
    exit 1
}
