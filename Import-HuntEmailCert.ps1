#Requires -Version 7.0
<#
.SYNOPSIS
    Compatibility shim that calls Import-HuntEmailCert from the huntEmail module.

.DESCRIPTION
    One-time PFX import. The function lives in huntEmail.psm1.
    A failed import exits 1 when this file is executed.

.EXAMPLE
    $pwd = Read-Host "PFX password" -AsSecureString
    .\Import-HuntEmailCert.ps1 -Password $pwd
#>
[CmdletBinding()]
param (
    [Parameter()]
    [string]$PfxPath = "C:\certs\huntEmail.pfx",

    [Parameter(Mandatory)]
    [SecureString]$Password
)

Import-Module -Name (Join-Path $PSScriptRoot 'huntEmail.psd1') -Force -ErrorAction Stop

try {
    Import-HuntEmailCert @PSBoundParameters -ErrorAction Stop
}
catch {
    if ($MyInvocation.InvocationName -eq '.') { throw }
    exit 1
}
