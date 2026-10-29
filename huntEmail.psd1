@{
    RootModule        = 'huntEmail.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = '80c28c0c-abfa-43ab-82bd-8aec7180f1de'
    Author            = 'Security Operations'
    CompanyName       = 'huntEmail contributors'
    Copyright         = '(c) 2026 huntEmail contributors'
    Description       = 'Hunt, timeline, and hard-delete Microsoft 365 phishing mail, and read Copilot interactions. One module.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Write-NyanProgress',
        'Write-NyanComplete',
        'Get-HuntEmailPagedCollection',
        'Get-HuntEmailFolderMap',
        'huntEmail',
        'timelineEmail',
        'purgeEmail',
        'Connect-HuntEmailGraph',
        'Import-HuntEmailCert',
        'huntCopilot',
        'auditCopilot',
        'Start-HuntEmail'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @('hunt', 'timeline', 'purge', 'hcop', 'acop')
    PrivateData       = @{
        PSData = @{
            Tags       = @('Microsoft365', 'Graph', 'Security', 'Email', 'Copilot')
            ProjectUri = 'https://github.com/blacknet101/huntEmail'
            LicenseUri = 'https://github.com/blacknet101/huntEmail/blob/main/LICENSE'
        }
    }
}
