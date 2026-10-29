#Requires -Version 7.0
<#
.SYNOPSIS
    Offline parse, export, and mocked-URI checks for the huntEmail module.
.DESCRIPTION
    Does not contact Microsoft Graph or Exchange Online. Graph and audit calls
    are satisfied by in-module test hooks.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Assert-True {
    param ([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERT: $Message" }
}

$parseTargets = @(
    'huntEmail.psm1',
    'huntEmail.ps1',
    'timelineEmail.ps1',
    'purgeEmail.ps1',
    'Connect-HuntEmailGraph.ps1',
    'Import-HuntEmailCert.ps1',
    'cert/createCert.ps1'
)

foreach ($relative in $parseTargets) {
    $path = Join-Path $root $relative
    $tokens = $null
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $detail = ($errors | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
        throw "Parse failed for ${relative}: $detail"
    }
    Write-Host "parsed $relative"
}

$manifestPath = Join-Path $root 'huntEmail.psd1'
$null = Test-ModuleManifest -Path $manifestPath
Import-Module $manifestPath -Force

$expected = [ordered]@{
    'huntEmail'                  = 'hunt'
    'timelineEmail'              = 'timeline'
    'purgeEmail'                 = 'purge'
    'huntCopilot'                = 'hcop'
    'auditCopilot'               = 'acop'
    'Connect-HuntEmailGraph'     = $null
    'Import-HuntEmailCert'       = $null
    'Start-HuntEmail'            = $null
    'Write-NyanProgress'         = $null
    'Write-NyanComplete'         = $null
    'Get-HuntEmailPagedCollection' = $null
    'Get-HuntEmailFolderMap'     = $null
}

foreach ($name in $expected.Keys) {
    $command = Get-Command $name -ErrorAction Stop
    Assert-True ($command.ModuleName -eq 'huntEmail') "$name did not load from the huntEmail module (source: $($command.Source))"
    $aliasName = $expected[$name]
    if ($aliasName) {
        $alias = Get-Alias $aliasName -ErrorAction Stop
        Assert-True ($alias.Definition -eq $name) "Alias $aliasName points at $($alias.Definition)"
        Assert-True ($alias.ModuleName -eq 'huntEmail') "Alias $aliasName did not load from the huntEmail module"
    }
}

Assert-True ((Get-Command timelineEmail).Parameters.ContainsKey('IncludeCopilot')) "timelineEmail is missing -IncludeCopilot"
Assert-True ((Get-Command purgeEmail).Parameters.ContainsKey('WhatIf')) "purgeEmail is missing -WhatIf"
Assert-True (-not (Get-Command -Name '*Copilot*' -Module huntEmail | Where-Object { $_.Name -match 'purge|delete|remove' })) "A Copilot purge command was exported"

$source = Get-Content -LiteralPath (Join-Path $root 'huntEmail.psm1') -Raw
Assert-True ($source -match 'https://graph.microsoft.com/v1.0/copilot/users/') "Copilot Graph URI is missing"
Assert-True ($source -match 'interactionHistory/getAllEnterpriseInteractions') "getAllEnterpriseInteractions is missing"
Assert-True ($source -match 'AiEnterpriseInteraction\.Read\.All') "Copilot permission is missing"
Assert-True ($source -match "RecordType\s+=\s+'CopilotInteraction'") "Audit RecordType is missing"
Assert-True ($source -match "Operations\s+=\s+'CopilotInteraction'") "Audit Operations filter is missing"
Assert-True ($source -match 'ReturnLargeSet') "SessionCommand ReturnLargeSet is missing"
Assert-True ($source -match 'permanentDelete') "purge permanentDelete URI is missing"
Assert-True ($source -match "Type 'PURGE'") "typed PURGE confirmation is missing"
Assert-True ($source -match "ConfirmImpact = 'High'") "purge ConfirmImpact is missing"
Assert-True ($source -notmatch 'purgeData') "The module must not call eDiscovery purgeData"
Assert-True ($source -match [regex]::Escape('https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048')) "Docs link is missing from the module"

$captured = [System.Collections.Generic.List[string]]::new()
$graphHook = {
    param($Method, $Uri, $Headers, $OutputFilePath)
    $null = $Headers
    $null = $OutputFilePath
    $captured.Add("$Method $Uri")
    if ($Uri -like '*getAllEnterpriseInteractions*') {
        return @{
            value = @(
                [PSCustomObject]@{
                    id               = '1732148356886'
                    sessionId        = '19:session@thread.v2'
                    requestId        = 'f128b7a9-657e-b420-201e-5602114a29f4'
                    appClass         = 'IPM.SkypeTeams.Message.Copilot.BizChat'
                    interactionType  = 'userPrompt'
                    conversationType = 'bizchat'
                    createdDateTime  = '2026-02-18T15:00:00Z'
                    locale           = 'en-us'
                    contexts         = @()
                    from             = @{ user = @{ displayName = 'Example User' } }
                    body             = @{ contentType = 'text'; content = 'What should be on my radar?' }
                    attachments      = @()
                    links            = @()
                }
            )
        }
    }
    if ($Uri -like '*mailFolders*' -or $Uri -like '*/attachments*') { return @{ value = @() } }
    if ($Uri -like '*/messages*' -and $Uri -notlike '*permanentDelete*') {
        return @{
            value = @(
                [PSCustomObject]@{
                    id                = 'AAMkAG-test'
                    subject           = 'Invoice'
                    from              = @{ emailAddress = @{ address = 'phish@bad.example' } }
                    toRecipients      = @(@{ emailAddress = @{ address = 'user@example.com' } })
                    ccRecipients      = @()
                    receivedDateTime  = '2026-02-18T15:00:00Z'
                    sentDateTime      = '2026-02-18T14:59:00Z'
                    isRead            = $false
                    hasAttachments    = $false
                    parentFolderId    = 'inbox'
                    internetMessageId = '<abc@example>'
                    bodyPreview       = 'hello'
                }
            )
        }
    }
    return @{ value = @() }
}.GetNewClosure()

$module = Get-Module huntEmail
$module.NewBoundScriptBlock({
    param($Hook)
    $script:HuntEmailInvokeGraph = $Hook
    function script:Get-MgContext { [PSCustomObject]@{ ClientId = 'test-client'; AuthType = 'Certificate' } }
}).Invoke($graphHook)

$huntRows = @(huntEmail -Sender 'phish@bad.example' -UserPrincipalName 'user@example.com' -StartDate '2026-02-18T00:00:00Z' -EndDate '2026-02-18T23:59:59Z')
Assert-True ($huntRows.Count -eq 1) "huntEmail should return the mocked message (got $($huntRows.Count))"
Assert-True ($huntRows[0].MessageId -eq 'AAMkAG-test') "huntEmail MessageId mismatch"
Assert-True ($huntRows[0].InternetMessageId -eq '<abc@example>') "huntEmail InternetMessageId mismatch"
$searchUri = @($captured | Where-Object { $_ -like '*messages?*' -and $_ -like '*$search*' })
Assert-True ($searchUri.Count -ge 1) "huntEmail did not call the messages `$search URI"
Assert-True ($searchUri[0] -match 'phish') "huntEmail search URI did not include the sender"

$captured.Clear()
$timeline = @($huntRows | timelineEmail -IncludeCopilot)
Assert-True ($timeline.Count -eq 1) "timelineEmail should return one row"
Assert-True ($timeline[0].CopilotUsed -eq $true) "IncludeCopilot should flag the mocked interaction"
Assert-True ($timeline[0].CopilotAppClasses -match 'BizChat') "IncludeCopilot should record the app class"
$copilotUri = @($captured | Where-Object { $_ -like '*getAllEnterpriseInteractions*' })
Assert-True ($copilotUri.Count -ge 1) "timeline -IncludeCopilot did not call interaction history"

$captured.Clear()
$direct = @(huntCopilot -UserIds 'user@example.com' -StartDate '2026-02-01T00:00:00Z' -EndDate '2026-02-18T23:59:59Z' -AppClass BizChat)
Assert-True ($direct.Count -eq 1) "huntCopilot should return one interaction"
Assert-True ($direct[0].Body -eq 'What should be on my radar?') "huntCopilot did not flatten the prompt body"
Assert-True ($direct[0].InteractionType -eq 'userPrompt') "huntCopilot interaction type mismatch"
Assert-True ($direct[0].AppClass -eq 'IPM.SkypeTeams.Message.Copilot.BizChat') "huntCopilot app class mismatch"
$directUri = [string]@($captured | Where-Object { $_ -like '*getAllEnterpriseInteractions*' })[0]
$decoded = [uri]::UnescapeDataString(($directUri -split ' ', 2)[1])
Assert-True ($decoded -match '^https://graph.microsoft.com/v1.0/copilot/users/user@example.com/interactionHistory/getAllEnterpriseInteractions') "huntCopilot URI was '$decoded'"
Assert-True ($decoded -match '\$top=100') "huntCopilot did not send the documented `$top=100"
Assert-True ($decoded -match "appClass eq 'IPM.SkypeTeams.Message.Copilot.BizChat'") "huntCopilot appClass filter missing"
Assert-True ($decoded -match 'createdDateTime gt ') "huntCopilot createdDateTime gt bound missing"
Assert-True ($decoded -match 'createdDateTime lt ') "huntCopilot createdDateTime lt bound missing"

$forbidden = {
    param($Method, $Uri, $Headers, $OutputFilePath)
    $null = $Method, $Uri, $Headers, $OutputFilePath
    $message = 'Response status code does not indicate success: 403 (Forbidden).'
    throw $message
}
$module.NewBoundScriptBlock({
    param($Hook)
    $script:HuntEmailInvokeGraph = $Hook
}).Invoke($forbidden)

$forbiddenError = $null
try {
    huntCopilot -UserIds 'user@example.com' -StartDate '2026-02-01' -EndDate '2026-02-02' -ErrorAction Stop | Out-Null
}
catch {
    $forbiddenError = $_.Exception.Message
}
Assert-True ($forbiddenError -match 'HUNTEMAIL_COPILOT_FORBIDDEN') "403 did not fail as a missing Copilot permission (got: $forbiddenError)"
Assert-True ($forbiddenError -match 'AiEnterpriseInteraction.Read.All') "403 message did not name AiEnterpriseInteraction.Read.All"

# Restore the successful Graph hook for purge -WhatIf and the audit pipeline.
$module.NewBoundScriptBlock({
    param($Hook)
    $script:HuntEmailInvokeGraph = $Hook
}).Invoke($graphHook)
$captured.Clear()

$purgeInput = [PSCustomObject]@{
    Mailbox   = 'user@example.com'
    MessageId = 'AAMkAG-test'
    Subject   = 'Invoice'
    From      = 'phish@bad.example'
}
$purgeLog = @($purgeInput | purgeEmail -WhatIf -Confirm:$false)
Assert-True ($purgeLog.Count -eq 1) "purge -WhatIf should return one log row"
Assert-True ($purgeLog[0].Status -eq 'Skipped') "purge -WhatIf should skip the delete (status $($purgeLog[0].Status))"
Assert-True (-not ($captured | Where-Object { $_ -like '*permanentDelete*' })) "purge -WhatIf issued a permanentDelete call: $($captured -join ' | ')"

$auditCalls = [System.Collections.Generic.List[object]]::new()
$auditJson = @'
{
  "CreationTime": "2026-02-18T15:04:00",
  "Id": "record-1",
  "Operation": "CopilotInteraction",
  "RecordType": "CopilotInteraction",
  "UserId": "user@example.com",
  "Workload": "Copilot",
  "ClientRegion": "US",
  "CopilotEventData": {
    "AppHost": "Word",
    "AppIdentity": "Copilot.MicrosoftCopilot.Microsoft365Copilot",
    "ThreadId": "19:thread@thread.v2",
    "Contexts": [{ "Id": "https://example.sharepoint.com/file.docx", "Type": "docx" }],
    "Messages": [
      { "Id": "m-prompt", "isPrompt": true },
      { "Id": "m-response", "isPrompt": false }
    ],
    "AccessedResources": [{
      "Action": "Read",
      "Id": "resource-1",
      "Name": "Document1.docx",
      "SensitivityLabelId": "label-1",
      "SiteUrl": "https://example.sharepoint.com/sites/test/Document1.docx",
      "Type": "docx",
      "listItemUniqueId": "item-1"
    }],
    "AISystemPlugin": [{ "Id": "BingWebSearch", "Name": "BuiltIn" }],
    "ModelTransparencyDetails": [{ "ModelProviderName": "ExampleProvider", "ModelName": "example-model" }]
  }
}
'@
$auditHook = {
    param($Parameters)
    $auditCalls.Add($Parameters)
    if ($auditCalls.Count -gt 1) { return @() }
    return @(
        [PSCustomObject]@{
            AuditData   = $auditJson
            ResultIndex = 1
            ResultCount = 1
            CreationDate = '2026-02-18T15:04:00'
            UserIds     = 'user@example.com'
        }
    )
}.GetNewClosure()

$module.NewBoundScriptBlock({
    param($AuditHook)
    $script:HuntEmailSearchAuditLog = $AuditHook
    $script:HuntEmailExchangeConnectionTest = { $true }
}).Invoke($auditHook)

# Prove the missing-cmdlet message before the hook is what production uses.
$module.NewBoundScriptBlock({
    $script:HuntEmailExchangeConnectionTest = $null
    $script:HuntEmailSearchAuditLog = $null
}).Invoke()

$missing = @()
auditCopilot -UserIds 'user@example.com' -StartDate '2026-02-01' -EndDate '2026-02-02' -ErrorVariable missing -ErrorAction SilentlyContinue | Out-Null
Assert-True (($missing | ForEach-Object { "$_" }) -join ' ' -match 'Connect-ExchangeOnline') "Missing Exchange session did not mention Connect-ExchangeOnline"

$module.NewBoundScriptBlock({
    param($AuditHook)
    $script:HuntEmailSearchAuditLog = $AuditHook
    $script:HuntEmailExchangeConnectionTest = { $true }
}).Invoke($auditHook)

$auditRows = @(auditCopilot -UserIds 'user@example.com' -StartDate '2026-02-01T00:00:00Z' -EndDate '2026-02-18T23:59:59Z')
Assert-True ($auditRows.Count -eq 1) "auditCopilot should flatten one audit record"
Assert-True ($auditRows[0].UserId -eq 'user@example.com') "audit UserId mismatch"
Assert-True ($auditRows[0].AppHost -eq 'Word') "audit AppHost mismatch"
Assert-True ($auditRows[0].AccessedResources -eq 'Document1.docx') "audit accessed resource name mismatch"
Assert-True ($auditRows[0].SensitivityLabelIds -eq 'label-1') "audit sensitivity label mismatch"
Assert-True ($auditRows[0].AccessedResourceUrls -match 'example.sharepoint.com') "audit SiteUrl mismatch"
Assert-True ($auditRows[0].PromptMessageIds -eq 'm-prompt') "audit prompt message id mismatch"
Assert-True ($auditRows[0].ResponseMessageIds -eq 'm-response') "audit response message id mismatch"
Assert-True ($auditRows[0].ContextTypes -eq 'docx') "audit context type mismatch"
Assert-True ($auditCalls.Count -ge 1) "Search-UnifiedAuditLog was not called"
$auditParams = $auditCalls[0]
Assert-True ($auditParams.RecordType -eq 'CopilotInteraction') "RecordType was $($auditParams.RecordType)"
Assert-True ($auditParams.Operations -eq 'CopilotInteraction') "Operations was $($auditParams.Operations)"
Assert-True ($auditParams.SessionCommand -eq 'ReturnLargeSet') "SessionCommand was $($auditParams.SessionCommand)"
Assert-True (-not [string]::IsNullOrWhiteSpace([string]$auditParams.SessionId)) "SessionId was empty"
Assert-True ($auditParams.ResultSize -eq 5000) "ResultSize was $($auditParams.ResultSize)"
Assert-True ($auditParams.Formatted -eq $true) "Formatted switch was not set"
Assert-True (@($auditParams.UserIds) -contains 'user@example.com') "UserIds were not passed to Search-UnifiedAuditLog"

$captured.Clear()
$piped = @($auditRows | huntCopilot -AppClass Word)
Assert-True ($piped.Count -eq 1) "auditCopilot | huntCopilot should return the interaction"
$pipedUri = [uri]::UnescapeDataString(([string]@($captured | Where-Object { $_ -like '*getAllEnterpriseInteractions*' })[0] -split ' ', 2)[1])
Assert-True ($pipedUri -match 'user@example.com') "Piped huntCopilot did not use the audit UserId: $pipedUri"
Assert-True ($pipedUri -match 'IPM.SkypeTeams.Message.Copilot.Word') "Piped huntCopilot did not honor -AppClass Word"

Write-Host "Verify-HuntEmailModule passed" -ForegroundColor Green
