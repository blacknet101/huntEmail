#Requires -Version 7.0
<#
.SYNOPSIS
    Compatibility shim. timelineEmail now lives in the huntEmail module.

.DESCRIPTION
    Dot-sourcing this file imports huntEmail.psd1 so existing
    ". .\timelineEmail.ps1" commands keep working. Prefer:
        Import-Module .\huntEmail.psd1
    or:
        . .\huntEmail.ps1
#>

Write-Warning "timelineEmail.ps1 is a compatibility shim. Import ./huntEmail.psd1 (or dot-source ./huntEmail.ps1) instead."
Import-Module -Name (Join-Path $PSScriptRoot 'huntEmail.psd1') -Force -ErrorAction Stop
