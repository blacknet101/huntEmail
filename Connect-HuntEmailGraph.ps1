#Requires -Version 7.0
<#
.SYNOPSIS
    Compatibility shim that calls Connect-HuntEmailGraph from the huntEmail module.

.DESCRIPTION
    Existing ".\Connect-HuntEmailGraph.ps1" commands still connect with certificate auth.
    The function itself lives in huntEmail.psm1. A failed connection exits 1 when this
    file is executed, matching the previous script.

.NOTES
    Prefer: Import-Module .\huntEmail.psd1; Connect-HuntEmailGraph
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

Import-Module -Name (Join-Path $PSScriptRoot 'huntEmail.psd1') -Force -ErrorAction Stop

try {
    Connect-HuntEmailGraph @PSBoundParameters -ErrorAction Stop
}
catch {
    if ($MyInvocation.InvocationName -eq '.') { throw }
    exit 1
}
