#Requires -Version 7.0
<#
.SYNOPSIS
    Load the huntEmail module. Running this file opens the numbered menu.

.DESCRIPTION
    Dot-source this file, or Import-Module ./huntEmail.psd1, to load every function
    without the menu. Direct function calls stay the primary interface.

    Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048
#>
[CmdletBinding()]
param ()

$manifest = Join-Path $PSScriptRoot 'huntEmail.psd1'
Import-Module -Name $manifest -Force -ErrorAction Stop

# InvocationName is '.' when the file is dot-sourced. Executing the file opens the menu.
if ($MyInvocation.InvocationName -ne '.') {
    Start-HuntEmail
}
