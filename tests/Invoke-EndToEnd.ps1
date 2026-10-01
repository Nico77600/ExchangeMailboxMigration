#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Mailbox Migration - end-to-end run on the fake Exchange organisation.

.DESCRIPTION
    Runs the whole lifecycle with the real entry script against tests\FakeExchange.ps1:
        Inventory -> Plan System -> Start System -> Plan User -> Start User -WhatIf -> Start User
        -> Status -> Complete (not synced: skipped) -> Status (synced) -> Complete -> Cleanup
    and checks the exit codes and the key rules (monitoring never moved, no wipe of other move requests,
    Synced batches with pending moves never removed...).

    Needs no Pester: it runs in Windows PowerShell 5.1 (the Exchange Management Shell) as in PowerShell 7.
    It also produces the sample reports used for the screenshots of the guide (-OutputPath, -KeepOutput).

.PARAMETER OutputPath
    Folder of the reports and logs. Default: a temporary folder, deleted at the end unless -KeepOutput.

.EXAMPLE
    powershell.exe -NoProfile -File .\tests\Invoke-EndToEnd.ps1
    pwsh -NoProfile -File .\tests\Invoke-EndToEnd.ps1 -OutputPath C:\Temp\emm-demo -KeepOutput

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Exit code: 0 = every check passed, 1 = at least one check failed.
#>
[CmdletBinding()]
param([string]$OutputPath, [switch]$KeepOutput, [switch]$ShowOutput)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$entry = Join-Path $root 'Invoke-ExchangeMailboxMigration.ps1'
. (Join-Path $PSScriptRoot 'FakeExchange.ps1')
if (-not $OutputPath) { $OutputPath = Join-Path ([IO.Path]::GetTempPath()) ('EmmEndToEnd-' + [guid]::NewGuid().ToString('N')) } else { $KeepOutput = $true }
[void][IO.Directory]::CreateDirectory($OutputPath)

# Test configuration: the delivered file, with local folders and short waits.
$config = [IO.File]::ReadAllText((Join-Path $root 'config\ExchangeMailboxMigration.config.psd1'))
$replace = [ordered]@{
    "(?m)^(\s*OutputPath\s*=\s*)'[^']*'"         = "`${1}'$((Join-Path $OutputPath 'reports').Replace("'", "''"))'"
    "(?m)^(\s*Path\s*=\s*)'\.\\logs'"            = "`${1}'$((Join-Path $OutputPath 'logs').Replace("'", "''"))'"
    "(?m)^(\s*PollSeconds\s*=\s*)\d+"            = '${1}1'
    "(?m)^(\s*SettleSeconds\s*=\s*)\d+"          = '${1}0'
    "(?m)^(\s*OpenReport\s*=\s*)\`$true"         = '${1}$false'
    "(?m)^(\s*BatchCount\s*=\s*)\d+"             = '${1}4'
    "(?m)^(\s*ExcludeMailboxes\s*=\s*)@\(\)"     = "`${1}@('ceo@contoso.com')"
    "(?m)^(\s*TargetDatabasePattern\s*=\s*)'[^']*'" = "`${1}'^DB-0[1-4]`$`$'"
    "(?m)^(\s*Organization\s*=\s*)''"            = "`${1}'Contoso - Exchange 2019 (fictitious)'"
}
foreach ($k in $replace.Keys) {
    if ([regex]::Matches($config, $k).Count -ne 1) { throw "Test configuration: pattern not found once: $k" }
    $config = [regex]::Replace($config, $k, $replace[$k])
}
$configPath = Join-Path $OutputPath 'test.config.psd1'
[IO.File]::WriteAllText($configPath, $config, (New-Object Text.UTF8Encoding($true)))

$script:failures = 0
function Check([string]$Name, [bool]$Condition, [string]$Detail = '') {
    if ($Condition) { Write-Host ("  [PASS] {0}" -f $Name) -ForegroundColor Green }
    else { $script:failures++; Write-Host ("  [FAIL] {0} {1}" -f $Name, $Detail) -ForegroundColor Red }
}
function Run([string]$Label, [object[]]$Arguments) {
    Write-Host ''
    Write-Host ("=== {0}: {1}" -f $Label, ($Arguments -join ' ')) -ForegroundColor Cyan
    $params = @{ ConfigPath = $configPath }
    for ($i = 0; $i -lt $Arguments.Count; $i++) {
        $name = ([string]$Arguments[$i]).TrimStart('-')
        if ($i + 1 -lt $Arguments.Count -and -not ([string]$Arguments[$i + 1]).StartsWith('-')) { $params[$name] = $Arguments[$i + 1]; $i++ }
        else { $params[$name] = $true }
    }
    $global:LASTEXITCODE = $null
    if ($ShowOutput) { & $entry @params } else { & $entry @params 6>$null | Out-Null }
    $code = $LASTEXITCODE
    Write-Host ("    exit code {0}" -f $code) -ForegroundColor DarkGray
    return $code
}
function Latest([string]$Pattern) { @(Get-ChildItem (Join-Path $OutputPath 'reports') -Directory | Where-Object { $_.Name -like $Pattern } | Sort-Object Name -Descending)[0] }

New-EmmFakeOrg
$global:EmmFake.AutoStep = $true
$monitoring = @($global:EmmFake.Mailboxes | Where-Object { $_.Kind -eq 'Monitoring' })
$monitoringDbBefore = @($monitoring | ForEach-Object { $_.Database }) -join ','
$helpdeskBusy = @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq 'Helpdesk-2026' })[0]

try {
    # ---- Inventory -------------------------------------------------------------------------------------------------
    $code = Run 'Inventory' @()
    Check 'Inventory exit code 0' ($code -eq 0) "($code)"
    $inv = Latest '*_Inventory_All'
    Check 'Inventory report written' ($inv -and (Test-Path (Join-Path $inv.FullName 'Inventory.html')) -and (Test-Path (Join-Path $inv.FullName 'Inventory-Mailboxes.csv')))
    $rows = Import-Csv (Join-Path $inv.FullName 'Inventory-Mailboxes.csv') -Delimiter ';'
    Check 'Monitoring mailboxes listed and excluded' (@($rows | Where-Object { $_.Category -eq 'Monitoring' -and $_.Status -eq 'Excluded' }).Count -eq $monitoring.Count)
    Check 'Excluded mailbox (ceo) not moved' (@($rows | Where-Object { $_.Alias -eq 'ceo' -and $_.Status -eq 'Excluded' }).Count -eq 1)
    Check 'Archive left on a source database: ArchiveOnly' (@($rows | Where-Object { $_.Alias -eq 'archive.left' -and $_.MoveType -eq 'ArchiveOnly' }).Count -eq 1)
    Check 'Archive already on a target: PrimaryOnly' (@($rows | Where-Object { $_.Alias -eq 'archive.moved' -and $_.MoveType -eq 'PrimaryOnly' }).Count -eq 1)
    Check 'Team mailbox not selected (type not configured)' (@($rows | Where-Object { $_.RecipientTypeDetails -eq 'TeamMailbox' -and $_.Status -eq 'NotSelected' }).Count -eq 1)
    Check 'Recovery database ignored' (-not (@(Import-Csv (Join-Path $inv.FullName 'Inventory-Databases.csv') -Delimiter ';') | Where-Object { $_.Name -eq 'RDB01' }))

    # ---- System workload -------------------------------------------------------------------------------------------
    $code = Run 'Plan System' @('-Mode', 'Plan', '-Workload', 'System')
    Check 'Plan System exit code 0' ($code -eq 0) "($code)"
    $code = Run 'Start System' @('-Mode', 'Start', '-Workload', 'System', '-Force')
    Check 'Start System exit code 0' ($code -eq 0) "($code)"
    $sysCalls = @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'New-MoveRequest' -and -not $_.WhatIf })
    Check 'System: one move request per system mailbox (8)' ($sysCalls.Count -eq 8) "($($sysCalls.Count))"
    Check 'System: BatchName label set' (-not @($sysCalls | Where-Object { $_.Parameters['BatchName'] -ne 'EMM-SystemMailboxes' }).Count)
    Check 'System: AllowLargeItems without LargeItemLimit' (-not @($sysCalls | Where-Object { -not $_.Parameters['AllowLargeItems'] -or $_.Parameters.ContainsKey('LargeItemLimit') }).Count)
    Check 'System: mailboxes now on target databases' (-not @($global:EmmFake.Mailboxes | Where-Object { $_.Kind -in 'Arbitration', 'AuditLog', 'AuxAuditLog' -and $_.Database -notmatch '^DB-0[1-4]$' }).Count)

    # ---- User workload -----------------------------------------------------------------------------------------------
    $code = Run 'Plan User' @('-Mode', 'Plan', '-Workload', 'User')
    Check 'Plan User exit code 0' ($code -eq 0) "($code)"
    $plan = Latest '*_Plan_User'
    $planRows = Import-Csv (Join-Path $plan.FullName 'MigrationPlan.csv') -Delimiter ';'
    Check 'Plan: 4 user batches' (@($planRows | ForEach-Object { $_.Batch } | Sort-Object -Unique).Count -eq 4)
    Check 'Plan: no monitoring mailbox' (-not @($planRows | Where-Object { $_.Name -like 'HealthMailbox*' }).Count)
    Check 'Plan: every target matches the target pattern' (-not @($planRows | Where-Object { ($_.TargetDatabase -and $_.TargetDatabase -notmatch '^DB-0[1-4]$') -or ($_.TargetArchiveDatabase -and $_.TargetArchiveDatabase -notmatch '^DB-0[1-4]$') }).Count)
    $counts = @($planRows | Group-Object Batch | ForEach-Object { $_.Count })
    Check 'Plan: balanced counts (max - min <= 1)' ((($counts | Measure-Object -Maximum).Maximum - ($counts | Measure-Object -Minimum).Minimum) -le 1)

    $callsBefore = $global:EmmFake.Calls.Count
    $code = Run 'Start User -WhatIf' @('-Mode', 'Start', '-Workload', 'User', '-WhatIf')
    Check 'Start -WhatIf exit code 2 (mailboxes skipped: move in progress, finished move of another origin)' ($code -eq 2) "($code)"
    $simCalls = @($global:EmmFake.Calls | Select-Object -Skip $callsBefore)
    Check 'Simulation: every change command sent with -WhatIf' ($simCalls.Count -gt 0 -and -not @($simCalls | Where-Object { -not $_.WhatIf }).Count)
    Check 'Simulation: nothing created' ($global:EmmFake.Batches.Count -eq 0)

    $code = Run 'Start User' @('-Mode', 'Start', '-Workload', 'User', '-Force')
    Check 'Start User exit code 2 (skipped mailbox)' ($code -eq 2) "($code)"
    Check 'Start: 4 batches created' ($global:EmmFake.Batches.Count -eq 4) "($($global:EmmFake.Batches.Count))"
    Check 'Start: move in progress of another tool untouched' ($global:EmmFake.MoveRequests.Contains($helpdeskBusy) -and $helpdeskBusy.Status -eq 'InProgress')
    $removed = @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'Remove-MoveRequest' -and -not $_.WhatIf })
    Check 'Start: only the finished move request of this tool removed (Batch09); another origin''s one kept' ($removed.Count -eq 1 -and $removed[0].Parameters['Identity'] -like '*Retry Me' -and @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq 'Helpdesk-2025' }).Count -eq 1)
    $csv = (@($global:EmmFake.Batches | ForEach-Object { $_.Rows }))
    Check 'Start: archive-only row (MailboxType ArchiveOnly, no TargetDatabase)' (@($csv | Where-Object { $_.EmailAddress -eq 'archive.left@contoso.com' -and $_.MailboxType -eq 'ArchiveOnly' -and -not $_.TargetDatabase -and $_.TargetArchiveDatabase }).Count -eq 1)
    Check 'Start: primary-only row' (@($csv | Where-Object { $_.EmailAddress -eq 'archive.moved@contoso.com' -and $_.MailboxType -eq 'PrimaryOnly' }).Count -eq 1)
    $pf = @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'New-MoveRequest' -and $_.Parameters['SuspendWhenReadyToComplete'] -and -not $_.WhatIf })
    Check 'Start: public folder mailboxes moved with New-MoveRequest, suspended, labelled with the batch' ($pf.Count -eq 2 -and -not @($pf | Where-Object { $_.Parameters['BatchName'] -notmatch '^Batch\d\d$' -or $_.Parameters.ContainsKey('PublicFolder') }).Count)

    $code = Run 'Start User again' @('-Mode', 'Start', '-Workload', 'User', '-Force')
    Check 'Start again: active batches skipped, nothing removed' ($global:EmmFake.Batches.Count -eq 4 -and @($global:EmmFake.Calls | Where-Object { $_.Command -in 'Remove-MigrationBatch', 'Remove-MoveRequest' -and -not $_.WhatIf }).Count -eq 1)

    # ---- Status and completion ------------------------------------------------------------------------------------------
    Step-EmmFakeOrg
    $code = Run 'Status' @('-Mode', 'Status')
    Check 'Status exit code 0 or 2' ($code -in 0, 2) "($code)"
    $code = Run 'Complete too early' @('-Mode', 'Complete', '-Batch', 'All', '-Force')
    Check 'Complete of a batch still syncing: skipped (exit 2)' ($code -eq 2 -and -not @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'Complete-MigrationBatch' }).Count)
    Step-EmmFakeOrg -Rounds 2
    $code = Run 'Status synced' @('-Mode', 'Status')
    Check 'Status synced: failed move reported (exit 2)' ($code -eq 2) "($code)"
    $st = Latest '*_Status'
    $stRows = Import-Csv (Join-Path $st.FullName 'Status-Mailboxes.csv') -Delimiter ';'
    Check 'Status: stalled move flagged' (@($stRows | Where-Object { $_.Quarantined -eq 'True' }).Count -ge 1)

    $code = Run 'Cleanup before completion' @('-Mode', 'Cleanup', '-Force')
    Check 'Cleanup never removes a Synced batch with pending moves' ($global:EmmFake.Batches.Count -eq 4 -and -not $global:EmmFake.ContainsKey('LostMoves'))

    $code = Run 'Complete scheduled' @('-Mode', 'Complete', '-Batch', '1', '-CompleteAfter', (Get-Date).AddHours(3).ToString('yyyy-MM-dd HH:mm'), '-Force')
    $sets = @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'Set-MoveRequest' })
    Check 'Scheduled completion: Set-MoveRequest -CompleteAfter on the moves of Batch01' ($sets.Count -ge 1 -and $code -eq 0) "($code, $($sets.Count))"

    $code = Run 'Complete' @('-Mode', 'Complete', '-Batch', '2,3,4', '-Force')
    Check 'Complete exit code 0' ($code -eq 0) "($code)"
    Check 'Complete: public folder moves resumed' (@($global:EmmFake.Calls | Where-Object { $_.Command -eq 'Resume-MoveRequest' }).Count -ge 2)
    Step-EmmFakeOrg -Rounds 2

    $code = Run 'Cleanup' @('-Mode', 'Cleanup', '-Force')
    Check 'Cleanup exit code 0' ($code -eq 0) "($code)"
    $failedBatch = (@($global:EmmFake.MoveRequests | Where-Object { $_.Status -eq 'Failed' })[0].BatchName -replace '^MigrationService:', '')
    $expectedLeft = @('Batch01', $failedBatch) | Sort-Object -Unique
    $left = @($global:EmmFake.Batches | ForEach-Object { $_.Identity } | Sort-Object)
    Check 'Cleanup: completed batches removed; scheduled Batch01 and the batch with a failed move kept' (($left -join ',') -eq ($expectedLeft -join ',')) "(left: $($left -join ','), expected: $($expectedLeft -join ','))"
    Check 'Cleanup: move requests of other tools untouched' (@($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -like 'Helpdesk-*' }).Count -eq 2)
    Check 'Cleanup: migration users of other origins untouched' (@($global:EmmFake.MigrationUsers | Where-Object { $_.Identity -in 'old.user@contoso.com', 'other.tool@contoso.com' }).Count -eq 2)
    Check 'Cleanup: no synchronised or failed move lost' (-not $global:EmmFake.ContainsKey('LostMoves'))
    $code = Run 'Cleanup -IncludeFailed' @('-Mode', 'Cleanup', '-IncludeFailed', '-Force')
    Check 'Cleanup -IncludeFailed: the batch with the failed move removed, scheduled Batch01 kept' ($code -eq 0 -and (@($global:EmmFake.Batches | ForEach-Object { $_.Identity }) -join ',') -eq 'Batch01' -and -not $global:EmmFake.ContainsKey('LostMoves')) "($code)"

    # ---- Parameter checks --------------------------------------------------------------------------------------------------
    $code = Run 'Complete without -Batch' @('-Mode', 'Complete', '-Force')
    Check 'Complete without -Batch refused (exit 1)' ($code -eq 1)
    $code = Run 'Plan with -Batch' @('-Mode', 'Plan', '-Batch', '1')
    Check '-Batch with -Mode Plan refused (exit 1)' ($code -eq 1)

    Check 'Monitoring mailboxes never moved' ((@($monitoring | ForEach-Object { $_.Database }) -join ',') -eq $monitoringDbBefore -and -not @($global:EmmFake.Calls | Where-Object { $_.Text -match 'HealthMailbox' }).Count)
}
finally {
    Write-Host ''
    if ($script:failures) { Write-Host ("  {0} check(s) FAILED" -f $script:failures) -ForegroundColor Red } else { Write-Host '  All checks passed' -ForegroundColor Green }
    Write-Host ("  PowerShell {0} - output: {1}" -f $PSVersionTable.PSVersion, $OutputPath) -ForegroundColor DarkGray
    if (-not $KeepOutput) { Remove-Item -LiteralPath $OutputPath -Recurse -Force -ErrorAction SilentlyContinue }
}
exit $(if ($script:failures) { 1 } else { 0 })
