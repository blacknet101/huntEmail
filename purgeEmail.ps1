#Requires -Version 7.0
<#
.SYNOPSIS
    Compatibility shim. purgeEmail now lives in the huntEmail module.

.DESCRIPTION
    Dot-sourcing this file imports huntEmail.psd1 so existing
    ". .\purgeEmail.ps1" commands keep working. Purge safety rails are unchanged.
    Prefer:
        Import-Module .\huntEmail.psd1
#>

Write-Warning "purgeEmail.ps1 is a compatibility shim. Import ./huntEmail.psd1 (or dot-source ./huntEmail.ps1) instead."
Import-Module -Name (Join-Path $PSScriptRoot 'huntEmail.psd1') -Force -ErrorAction Stop
