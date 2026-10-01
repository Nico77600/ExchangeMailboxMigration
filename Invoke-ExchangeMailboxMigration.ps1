#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Mailbox Migration - moves Exchange Server mailboxes from source databases to target
    databases with migration batches. System mailboxes and user mailboxes are kept apart; monitoring
    mailboxes are never moved.

.DESCRIPTION
    One script, one mode per step of a migration:

      Inventory  (default) Reads the databases and the mailboxes, shows what the selection would move.
                           Changes nothing.
      Plan                 Builds the migration plan: target database(s) and batch of every mailbox.
                           Changes nothing; writes MigrationPlan.csv / .json and an HTML report.
      Start                Submits the latest plan (or -PlanPath): system mailboxes with individual move
                           requests (completed automatically), user mailboxes with local migration batches
                           (stopped when synchronised). Asks for confirmation.
      Status               Real state of the batches and move requests. -Follow refreshes it until the end.
      Complete             Completes batches now, or at -CompleteAfter. Asks for confirmation.
      Cleanup              Removes the finished batches, move requests and migration users of this tool.
                           Asks for confirmation.

    Everything is set in config\ExchangeMailboxMigration.config.psd1; the command line chooses the mode,
    the workload and, for one execution, a few overrides. -WhatIf simulates Start, Complete and Cleanup:
    every Exchange command is sent with -WhatIf (Exchange checks it, nothing changes).

    Workloads:
      System  arbitration, audit log, auxiliary audit log and discovery mailboxes
      User    user, shared, room, equipment, linked and public folder mailboxes, and archives
      All     both (default of the configuration file)
    Monitoring mailboxes (HealthMailbox...) are never moved, whatever the configuration.

.PARAMETER Mode
    Inventory | Plan | Start | Status | Complete | Cleanup. Default: Inventory.

.PARAMETER Workload
    System | User | All. Overrides Scope.Workload. With Status and Cleanup: restricts to the system
    move requests or to the user batches.

.PARAMETER MailboxType
    Inventory and Plan: restricts this execution to these types, for example SharedMailbox,RoomMailbox
    or Archive (archives only). The workload is deduced from them when -Workload is not given.

.PARAMETER Batch
    Status, Complete and Cleanup: batches to act on. Numbers (1, 03), names (Batch03), System (the
    system move requests) or All. Required with Complete.

.PARAMETER CompleteAfter
    Complete: completion time instead of now, as 'yyyy-MM-dd HH:mm' (local time of this computer).

.PARAMETER Follow
    Status and Complete: refreshes the status every Status.RefreshMinutes until the selected batches are
    finished (Ctrl+C stops following, not the migration). The HTML report reloads itself.

.PARAMETER BatchCount
    Plan: number of batches (Balanced strategy). Overrides Plan.BatchCount.

.PARAMETER PlanPath
    Start: plan folder (or its MigrationPlan.json). Default: the latest plan of the output folder that
    covers the workload.

.PARAMETER IncludeFailed
    Cleanup: also removes failed batches and failed move requests (overrides Cleanup.IncludeFailed).
    Needed before a failed mailbox can be planned again.

.PARAMETER ConfigPath
    Configuration file. Default: config\ExchangeMailboxMigration.config.psd1 next to this script.

.PARAMETER OutputPath
    Overrides Report.OutputPath for this execution (plans are looked for there too).

.PARAMETER ExchangeServer
    Overrides Connection.ExchangeServer for this execution.

.PARAMETER Credential
    Account used for remote PowerShell (default: the current Windows account).

.PARAMETER Force
    Start, Complete and Cleanup: no confirmation question (scheduled task). Start also accepts a plan
    older than Plan.MaxPlanAgeDays.

.EXAMPLE
    .\Invoke-ExchangeMailboxMigration.ps1
    Inventory of the configured selection: nothing is changed.

.EXAMPLE
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan -Workload System
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System -WhatIf
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System
    System mailboxes first: plan, simulation, then real start.

.EXAMPLE
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan -Workload User -BatchCount 8
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload User
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Status -Follow
    User mailboxes and archives in 8 balanced batches, then follow-up.

.EXAMPLE
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Complete -Batch 01,02 -CompleteAfter '2026-10-03 22:00'
    Schedules the completion of Batch01 and Batch02 for the evening.

.EXAMPLE
    .\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan -MailboxType SharedMailbox,RoomMailbox,EquipmentMailbox
    A plan with the shared and resource mailboxes only.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Requires : Windows PowerShell 5.1 (Exchange Management Shell, or powershell.exe with remote PowerShell to
               an Exchange server). PowerShell 7 is not supported by Microsoft for Exchange Server management:
               the script stops with an explicit message when started in it.
    Exit codes : 0 = success, 1 = failure, 2 = finished with items to look at: Start skipped or failed
                 mailboxes, Complete skipped or failed batches, Status failed moves, Cleanup failed removals
                 (batches kept on purpose are listed but do not change the code), 3 = cancelled at the
                 confirmation question (nothing changed).
    Documentation : docs\ExchangeMailboxMigration-Guide.md (or .html)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Inventory', 'Plan', 'Start', 'Status', 'Complete', 'Cleanup')]
    [string]$Mode = 'Inventory',

    [ValidateSet('System', 'User', 'All')]
    [string]$Workload,
    [string[]]$MailboxType,
    [string[]]$Batch,
    [datetime]$CompleteAfter,
    [switch]$Follow,
    [ValidateRange(1, 99)]
    [int]$BatchCount,
    [string]$PlanPath,
    [switch]$IncludeFailed,

    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\ExchangeMailboxMigration.config.psd1'),
    [string]$OutputPath,
    [string]$ExchangeServer,
    [pscredential]$Credential,
    [switch]$Force
)

# -WhatIf is the simulation of the tool: it is passed explicitly to the Exchange commands. It is then
# switched off for this script, so the report and log files are written in simulation too.
$simulate = [bool]$WhatIfPreference
$WhatIfPreference = $false
# Exchange Server is managed with Windows PowerShell 5.1 only (Exchange Management Shell): PowerShell 7 is not
# supported by Microsoft for Exchange Server, so the tool refuses it rather than run in an unsupported way.
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Host ''
    Write-Host '  [ERROR] Exchange Mailbox Migration runs in Windows PowerShell 5.1 only (Exchange Management Shell).' -ForegroundColor Red
    Write-Host "          PowerShell $($PSVersionTable.PSVersion) is not supported by Microsoft for Exchange Server management." -ForegroundColor Red
    Write-Host '          Start it with powershell.exe, or from the Exchange Management Shell.' -ForegroundColor Red
    Write-Host ''
    exit 1
}
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
# Numbers and dates are displayed the same way on every server (1,234.5), whatever the regional settings.
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$clock = [Diagnostics.Stopwatch]::StartNew()
$exitCode = 1
$logPath = $null
$tool = '.\Invoke-ExchangeMailboxMigration.ps1'
$dot = [char]0x00B7

function Get-RunDuration { Format-EmmDuration $clock.Elapsed.TotalSeconds }

function Connect-Step {
    <# Step "Connecting to Exchange", shared by every mode. #>
    param([int]$Number, [int]$Total)
    Write-EmmStep $Number $Total 'Connecting to Exchange' -Icon Key
    $c = Connect-EmmExchange -Server $settings.Connection.ExchangeServer -Credential $Credential -Authentication $settings.Connection.Authentication
    Write-EmmItem Ok ('{0}  {1} {2}' -f $c.Account, $dot, $c.Method)
    return $c
}

function Get-ActionCount {
    <# Counts of the actions recorded in this execution, by status. #>
    $a = @(Get-EmmActions)
    return [pscustomobject]@{
        All = $a.Count; Success = @($a | Where-Object { $_.Status -eq 'Success' }).Count; Simulated = @($a | Where-Object { $_.Status -eq 'Simulated' }).Count
        Skipped = @($a | Where-Object { $_.Status -eq 'Skipped' }).Count; Failed = @($a | Where-Object { $_.Status -eq 'Failed' }).Count
        AlreadyDone = @($a | Where-Object { $_.Status -eq 'AlreadyDone' }).Count
    }
}

function Confirm-Change {
    <# Confirmation question before a change (not in simulation, not with -Force). Returns $false when declined. #>
    param([Parameter(Mandatory)][string]$Question)
    if ($simulate) { Write-EmmItem Info 'Simulation: every Exchange command is sent with -WhatIf, nothing changes.' -Icon Simulate; return $true }
    if ($Force) { Write-EmmItem Info '-Force: no confirmation question.' -Icon Next; return $true }
    try { $answer = $PSCmdlet.ShouldContinue($Question, 'Exchange Mailbox Migration') }
    catch { throw "Confirmation impossible in this session ($($_.Exception.Message)). Run interactively, or use -Force for a scheduled task." }
    if ($answer) { Write-EmmLog 'INFO' "Confirmed: $Question" } else { Write-EmmLog 'WARN' "Declined: $Question" }
    return $answer
}

function Write-Files {
    param([object[]]$Files)
    foreach ($f in $Files) {
        $rows = if ($null -ne $f.Rows) { "   $(Format-EmmNumber $f.Rows) row(s) $dot " } else { '   ' }
        Write-EmmItem Ok ('{0,-4} {1}{2}{3}' -f $f.Kind, (Split-Path $f.Path -Leaf), $rows, ('{0:N0} KB' -f [Math]::Ceiling($f.Bytes / 1KB))) -Icon File
    }
}

function Write-StatusTable {
    param([object[]]$Batches)
    $rows = foreach ($b in $Batches) {
        [pscustomobject]@{
            Status = $(if ($b.Failed) { 'Warn' } elseif ($b.Mailboxes -and $b.Completed -eq $b.Mailboxes) { 'Ok' } else { 'Info' })
            Batch = $b.Name; State = $b.Status; Mailboxes = Format-EmmNumber $b.Mailboxes; Done = $b.Completed; Ready = $b.Ready; Active = $b.InProgress; Failed = $b.Failed
            Progress = ('{0} {1,3}%' -f (Get-EmmBar $b.Percent 20), $b.Percent)
        }
    }
    Write-EmmTable -Columns ([ordered]@{ Batch = -20; State = -14; Mailboxes = 9; Done = 6; Ready = 6; Active = 6; Failed = 6; Progress = -27 }) -Rows @($rows)
}

function Get-StatusNext {
    <# Next command suggested by a status: complete what is ready, clean what is finished, otherwise follow. #>
    param($Status)
    $prefix = '^' + [regex]::Escape($settings.Plan.BatchNamePrefix) + '0*'
    $ready = @($Status.Batches | Where-Object { $_.Next -like '-Mode Complete*' } | ForEach-Object { if ($_.Name -eq $settings.System.BatchName) { 'System' } else { $n = $_.Name -replace $prefix, ''; if ($n -match '^\d+$') { $n.PadLeft(2, '0') } else { $_.Name } } })
    if ($ready.Count) { return @("$tool -Mode Complete -Batch $($ready -join ',')", 'Ready to complete (synchronised): complete now, or at a time chosen with -CompleteAfter.') }
    $all = @($Status.Batches)
    if ($all.Count -and -not @($all | Where-Object { $_.Completed -lt $_.Mailboxes }).Count) { return @("$tool -Mode Cleanup", 'Every move is completed: remove the finished batches and move requests.') }
    if ($all.Count) { return @("$tool -Mode Status -Follow", 'Follow the synchronisation until the batches are ready to complete.') }
    return @('', '')
}

function Invoke-StatusRound {
    <# Reads the status, writes the report (same file name in follow mode) and returns the status. #>
    param([string[]]$Filter, [string]$Directory, [int]$RefreshSeconds, [string]$FilterLabel)
    $status = Get-EmmStatus -Settings $settings -BatchName $Filter
    $next = Get-StatusNext $status
    $meta = [ordered]@{ workload = $(if ($Workload) { $Workload } else { 'All' }); connection = $connection.Method; account = $connection.Account; batchFilter = $FilterLabel; next = $next[0]; nextText = $next[1]; duration = (Get-RunDuration) }
    $batchColumns = @('Name', 'Workload', 'Status', 'Mailboxes', 'Completed', 'Ready', 'InProgress', 'Failed', 'Quarantined', 'Percent', 'SizeMB', 'TransferredMB', 'Next')
    $files = New-EmmReport -Kind Status -Settings $settings -Directory $Directory -Meta $meta -HtmlName 'Status.html' -RefreshSeconds $RefreshSeconds `
        -Tables ([ordered]@{ batches = $status.Batches; mailboxes = $status.Rows }) -Columns @{ batches = $batchColumns } `
        -Csv ([ordered]@{ 'Status-Batches.csv' = 'batches'; 'Status-Mailboxes.csv' = 'mailboxes' })
    return [pscustomobject]@{ Status = $status; Files = $files; Next = $next }
}

function Invoke-Follow {
    <# Follow mode: one status round every Status.RefreshMinutes until the selected batches are finished. #>
    param([string[]]$Filter, [string]$Directory, [string]$FilterLabel)
    $minutes = $settings.Status.RefreshMinutes
    $round = Invoke-StatusRound -Filter $Filter -Directory $Directory -RefreshSeconds ($minutes * 60) -FilterLabel $FilterLabel
    $html = @($round.Files | Where-Object { $_.Kind -eq 'HTML' })[0].Path
    if ($settings.Status.OpenReport) { try { Invoke-Item -LiteralPath $html } catch { Write-EmmItem Warn "The report could not be opened: $($_.Exception.Message)" } }
    Write-EmmItem Info ("Report refreshed every {0} min: {1}" -f $minutes, $html) -Icon Report
    Write-EmmItem Info 'Ctrl+C stops following (the migration continues).' -Icon Clock
    $iteration = 1
    while ($true) {
        $s = $round.Status
        $rows = @($s.Rows)
        $done = @($rows | Where-Object { $_.Group -eq 'Completed' }).Count; $ready = @($rows | Where-Object { $_.Group -eq 'Ready' }).Count
        $active = @($rows | Where-Object { $_.Group -eq 'InProgress' }).Count; $failed = @($rows | Where-Object { $_.Group -eq 'Failed' }).Count
        $avg = if ($rows.Count) { [int](@($rows | Measure-Object Percent -Average)[0].Average) } else { 0 }
        Write-EmmItem $(if ($failed) { 'Warn' } else { 'Info' }) ('{0}  {1} {2,3}%  completed {3}/{4} {5} ready {6} {5} in progress {7} {5} failed {8}' -f (Get-Date).ToString('HH:mm'), (Get-EmmBar $avg 16), $avg, $done, $rows.Count, $dot, $ready, $active, $failed) -Icon Sync
        $open = @($s.Batches | Where-Object { $_.Mailboxes -and ($_.Completed + $_.Failed) -lt $_.Mailboxes })
        if (-not $s.Batches.Count -or -not $open.Count) {
            Write-EmmItem Ok 'The selected batches are finished: follow-up stopped.' -Icon Flag
            $round = Invoke-StatusRound -Filter $Filter -Directory $Directory -RefreshSeconds 0 -FilterLabel $FilterLabel
            return $round
        }
        Start-Sleep -Seconds ($minutes * 60)
        $iteration++
        try { $round = Invoke-StatusRound -Filter $Filter -Directory $Directory -RefreshSeconds ($minutes * 60) -FilterLabel $FilterLabel }
        catch { Write-EmmItem Warn ("Round {0}: {1} - next try in {2} min" -f $iteration, $_.Exception.Message, $minutes) }
    }
}

try {
    # do { } while ($false): 'break' ends the execution early; the finally block always runs.
    do {
        Import-Module (Join-Path $PSScriptRoot 'ExchangeMailboxMigration.psd1') -Force

        # -----------------------------------------------------------------------------------------
        # Configuration file, command-line overrides and parameter checks.
        # -----------------------------------------------------------------------------------------
        $settings = Import-EmmConfiguration -Path $ConfigPath -Root $PSScriptRoot
        if ($OutputPath) {
            $full = if ([IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path (Get-Location).Path $OutputPath }
            $settings.Report.OutputPath = [IO.Path]::GetFullPath($full)
        }
        if ($ExchangeServer) { $settings.Connection.ExchangeServer = $ExchangeServer }
        if ($IncludeFailed) { $settings.Cleanup.IncludeFailed = $true }
        $allowed = @{
            MailboxType = 'Inventory', 'Plan'; BatchCount = 'Plan'; PlanPath = 'Start'; Batch = 'Status', 'Complete', 'Cleanup'
            CompleteAfter = 'Complete'; Follow = 'Status', 'Complete'; IncludeFailed = 'Cleanup'
        }
        foreach ($p in $allowed.Keys) {
            if ($PSBoundParameters.ContainsKey($p) -and $Mode -notin $allowed[$p]) { throw "-$p is used with -Mode $($allowed[$p] -join ' or '), not with -Mode $Mode." }
        }
        if ($Mode -eq 'Complete' -and -not $Batch) { throw '-Mode Complete needs -Batch (numbers, names, System or All): completion is always an explicit choice.' }
        if ($PSBoundParameters.ContainsKey('CompleteAfter') -and $CompleteAfter -le (Get-Date)) { throw "-CompleteAfter $($CompleteAfter.ToString('yyyy-MM-dd HH:mm')) is in the past. Omit it to complete now." }
        if ($simulate -and $Mode -in 'Inventory', 'Plan', 'Status') { Write-Warning "-WhatIf has no effect with -Mode $Mode, which never changes anything." }

        $logPath = Start-EmmLog -Directory $settings.Logging.Path -RetentionDays $settings.Logging.RetentionDays
        Reset-EmmActions
        $scope = $null
        if ($Mode -in 'Inventory', 'Plan') { $scope = Resolve-EmmScope -Settings $settings -Workload $Workload -MailboxType $MailboxType }

        $d = $settings.Databases
        $dbText = @(
            $(if ($d.SourceDatabasePattern) { $d.SourceDatabasePattern } else { '' }),
            $(if ($d.DatabaseMap.Count) { "map of $($d.DatabaseMap.Count)" } else { '' })
        ) | Where-Object { $_ }
        $banner = [ordered]@{ Mode = @($(if ($simulate) { 'Simulate' } else { 'Info' }), ($Mode + $(if ($simulate) { " $dot simulation (-WhatIf)" } else { '' }))) }
        if ($scope) { $banner['Selection'] = @($(if ($scope.Workload -eq 'System') { 'System' } else { 'People' }), $scope.Label) }
        elseif ($Workload) { $banner['Workload'] = @('People', $Workload) }
        if ($Batch) { $banner['Batches'] = @('Batch', ($Batch -join ', ')) }
        if ($PSBoundParameters.ContainsKey('CompleteAfter')) { $banner['Completion'] = @('Calendar', $CompleteAfter.ToString('yyyy-MM-dd HH:mm')) }
        $banner['Databases'] = @('Database', ('{0}  {1}  {2}' -f ($dbText -join ' + '), [char]0x2192, $(if ($d.TargetDatabasePattern) { $d.TargetDatabasePattern } else { 'map' })))
        $banner['Output'] = @('Folder', $settings.Report.OutputPath)
        $banner['Log'] = @('Log', $logPath)
        Write-EmmBanner -Title 'Exchange Mailbox Migration' -Subtitle "Exchange Server mailbox moves $dot system and user mailboxes kept apart" -Details $banner
        $connection = $null

        switch ($Mode) {
            # =====================================================================================
            # Inventory / Plan: read the organisation, decide, write (plan and) report. No change.
            # =====================================================================================
            { $_ -in 'Inventory', 'Plan' } {
                $total = 5
                $connection = Connect-Step 1 $total

                Write-EmmStep 2 $total 'Reading the mailbox databases' -Icon Database
                $dbInventory = Get-EmmDatabaseInventory -Settings $settings
                $shown = @($dbInventory.Databases | Where-Object { $_.Role -ne 'Other' })
                $other = $dbInventory.Databases.Count - $shown.Count
                Write-EmmTable -Columns ([ordered]@{ Database = -22; Role = -14; Server = -16; Mounted = -8; Size = 10 }) -Rows @($shown | ForEach-Object {
                        [pscustomobject]@{ Status = $(if ($_.Mounted -eq $false) { 'Warn' } elseif ($_.Role -eq 'Source') { 'Info' } else { 'Ok' }); Database = $_.Name; Role = $_.Role; Server = $_.Server; Mounted = $(if ($null -eq $_.Mounted) { '?' } elseif ($_.Mounted) { 'yes' } else { 'NO' }); Size = Format-EmmSize $_.SizeMB }
                    })
                Write-EmmItem Info ('{0} source {1} {2} target {1} {3} other database(s) (not shown)' -f $dbInventory.Classification.Sources.Count, $dot, @($dbInventory.Databases | Where-Object { $_.IsTarget -or $_.IsArchiveTarget }).Count, $other)
                foreach ($w in $dbInventory.Warnings) { Write-EmmItem Warn $w }
                if (-not $dbInventory.Classification.Sources.Count) { throw 'No source database: nothing can be moved. Check the Databases section of the configuration.' }

                Write-EmmStep 3 $total 'Reading the mailboxes' -Icon Mailbox
                $mbxInventory = Get-EmmMailboxInventory -Settings $settings -DatabaseInventory $dbInventory -Scope $scope
                $all = @($mbxInventory.Mailboxes)
                $toMove = @($all | Where-Object { $_.Status -eq 'ToMove' })
                $volume = [double](@($toMove | Measure-Object MoveSizeMB -Sum)[0].Sum)
                $count = { param($s) @($all | Where-Object { $_.Status -eq $s }).Count }
                Write-EmmItem Ok ('To move       {0,8}  {1} {2} {3} system {2} {4} user {2} {5} archive(s)' -f (Format-EmmNumber $toMove.Count), (Format-EmmSize $volume), $dot,
                    @($toMove | Where-Object { $_.Category -eq 'System' }).Count, @($toMove | Where-Object { $_.Category -eq 'User' }).Count, @($toMove | Where-Object { Test-EmmMoveType $_.MoveType 'Archive' }).Count) -Icon Move
                Write-EmmItem Info ('On target     {0,8}' -f (Format-EmmNumber (& $count 'OnTarget'))) -Icon Target
                $notSelected = & $count 'NotSelected'
                Write-EmmItem $(if ($notSelected) { 'Info' } else { 'Skip' }) ('Not selected  {0,8}  on source databases, outside this selection' -f (Format-EmmNumber $notSelected)) -Icon Skip
                Write-EmmItem Info ('Excluded      {0,8}  of which {1} monitoring (never moved)' -f (Format-EmmNumber (& $count 'Excluded')), @($all | Where-Object { $_.Category -eq 'Monitoring' }).Count) -Icon Lock
                Write-EmmItem Info ('Outside       {0,8}  not on a source database' -f (Format-EmmNumber (& $count 'Outside'))) -Icon Info
                foreach ($w in $mbxInventory.Warnings) { Write-EmmItem Warn $w }
                $warnings = @($dbInventory.Warnings) + @($mbxInventory.Warnings)

                $inventoryColumns = @('Name', 'DisplayName', 'Alias', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'Category', 'Status', 'Reason', 'MoveType', 'Database', 'PrimaryRole',
                    'HasArchive', 'ArchiveDatabase', 'ArchiveRole', 'PrimarySizeMB', 'ArchiveSizeMB', 'MoveSizeMB', 'ItemCount', 'ArchiveItemCount', 'Server', 'Guid', 'ExchangeGuid', 'ArchiveGuid', 'DistinguishedName')
                $dbColumns = @('Name', 'Role', 'Server', 'Version', 'Mounted', 'SizeMB', 'WhitespaceMB', 'Mailboxes', 'Archives', 'SystemMailboxes', 'MonitoringMailboxes', 'ToMove', 'ToMoveMB', 'MapTarget')
                $baseMeta = [ordered]@{ workload = $scope.Workload; selection = $scope.Label; connection = $connection.Method; account = $connection.Account }

                if ($Mode -eq 'Inventory') {
                    $objects = Get-EmmMigrationObjects -Settings $settings
                    Write-EmmStep 4 $total 'Existing migration batches and move requests' -Icon Batch
                    $active = @($objects.ToolBatches | Where-Object { $_.Status -notin 'Completed', 'CompletedWithErrors' })
                    Write-EmmItem $(if ($active.Count) { 'Warn' } else { 'Ok' }) ('{0} batch(es) of this tool ({1} active) {2} {3} move request(s) of this tool' -f $objects.ToolBatches.Count, $active.Count, $dot, $objects.ToolMoves.Count) -Icon Batch
                    foreach ($b in $objects.ToolBatches) { Write-EmmItem Info ('{0,-20} {1,-20} {2} mailbox(es)' -f $b.Name, $b.Status, $b.Total) -Icon Batch }
                    if ($objects.OtherBatches -or $objects.OtherMoves) { Write-EmmItem Info ('Not managed by this tool: {0} batch(es), {1} move request(s) (never changed)' -f $objects.OtherBatches, $objects.OtherMoves) }

                    Write-EmmStep 5 $total 'Writing the report' -Icon Report
                    $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode 'Inventory' -Suffix $scope.Workload
                    $meta = $baseMeta; $meta['warnings'] = @($warnings); $meta['duration'] = Get-RunDuration
                    $meta['next'] = "$tool -Mode Plan -Workload $($scope.Workload)"; $meta['nextText'] = 'Build the plan of this selection: target databases and batches (changes nothing).'
                    $files = New-EmmReport -Kind Inventory -Settings $settings -Directory $dir -Meta $meta `
                        -Tables ([ordered]@{ databases = $dbInventory.Databases; mailboxes = $all }) -Columns @{ databases = $dbColumns; mailboxes = $inventoryColumns } `
                        -Csv ([ordered]@{ 'Inventory-Databases.csv' = 'databases'; 'Inventory-Mailboxes.csv' = 'mailboxes' })
                    Write-Files $files
                    Write-EmmSummary -Title 'Inventory ready' -Status $(if ($warnings.Count) { 'Warn' } else { 'Ok' }) -Values ([ordered]@{
                            'To move'   = @('Move', ('{0} mailbox(es) {1} {2}' -f (Format-EmmNumber $toMove.Count), $dot, (Format-EmmSize $volume)))
                            'Not moved' = @('Lock', ('{0} not selected {1} {2} excluded (monitoring included)' -f $notSelected, $dot, (& $count 'Excluded')))
                            'Report'    = @('Folder', $dir)
                            'Next'      = @('Next', "$tool -Mode Plan -Workload $($scope.Workload)")
                            'Duration'  = @('Clock', (Get-RunDuration))
                        })
                    $exitCode = 0
                    break
                }

                Write-EmmStep 4 $total 'Building the plan' -Icon Plan
                $plan = New-EmmPlan -Settings $settings -DatabaseInventory $dbInventory -Mailbox $all -Scope $scope -BatchCount $BatchCount
                if (-not $plan.Rows.Count) {
                    Write-EmmItem Warn 'Nothing to move for this selection.'
                } else {
                    Write-EmmTable -Columns ([ordered]@{ Batch = -22; Mailboxes = 9; Size = 10; Archives = 8; Types = -32 }) -Rows @($plan.Batches | ForEach-Object {
                            [pscustomobject]@{ Status = 'Ok'; Batch = $_.Name; Mailboxes = Format-EmmNumber $_.Mailboxes; Size = Format-EmmSize $_.SizeMB; Archives = $_.Archives; Types = $_.Types }
                        })
                    $targets = @($plan.Databases | Where-Object { $_.PlannedPrimaries + $_.PlannedArchives -gt 0 })
                    Write-EmmItem Info ('Targets: {0}' -f (@($targets | ForEach-Object { '{0} +{1}' -f $_.Name, (Format-EmmSize $_.PlannedMB) }) -join ', ')) -Icon Target
                }
                foreach ($w in $plan.Warnings) { Write-EmmItem Warn $w }

                Write-EmmStep 5 $total 'Writing the plan and the report' -Icon Report
                $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode 'Plan' -Suffix $scope.Workload
                foreach ($p in (Save-EmmPlan -Plan $plan -Directory $dir -Settings $settings)) { Write-EmmItem Ok ('PLAN {0}' -f (Split-Path $p -Leaf)) -Icon File }
                $meta = $baseMeta; $meta['planId'] = $plan.Meta.PlanId; $meta['created'] = $plan.Meta.Created; $meta['batchStrategy'] = $plan.Meta.BatchStrategy
                $meta['warnings'] = @($warnings) + @($plan.Warnings); $meta['duration'] = Get-RunDuration
                $meta['next'] = "$tool -Mode Start -Workload $($scope.Workload) -WhatIf"; $meta['nextText'] = 'Simulate the start of this plan (Exchange checks every command), then run the same command without -WhatIf.'
                $files = New-EmmReport -Kind Plan -Settings $settings -Directory $dir -Meta $meta `
                    -Tables ([ordered]@{ mailboxes = $plan.Rows; batches = $plan.Batches; databases = $plan.Databases }) `
                    -Csv ([ordered]@{ 'Plan-Batches.csv' = 'batches'; 'Plan-Databases.csv' = 'databases' })
                Write-Files $files
                Write-EmmSummary -Title $(if ($plan.Rows.Count) { 'Plan ready' } else { 'Nothing to plan' }) -Status $(if ($plan.Warnings.Count -or $warnings.Count) { 'Warn' } else { 'Ok' }) -Values ([ordered]@{
                        'Plan'      = @('Plan', ('{0} {1} {2}' -f $plan.Meta.PlanId, $dot, $scope.Label))
                        'Mailboxes' = @('Mailbox', ('{0} {1} {2} {1} {3} system {1} {4} user in {5} batch(es)' -f (Format-EmmNumber $plan.Meta.Mailboxes), $dot, (Format-EmmSize $plan.Meta.SizeMB), $plan.Meta.SystemMailboxes, $plan.Meta.UserMailboxes, $plan.Meta.BatchCount))
                        'Folder'    = @('Folder', $dir)
                        'Next'      = @('Next', "$tool -Mode Start -Workload $($scope.Workload) -WhatIf")
                        'Duration'  = @('Clock', (Get-RunDuration))
                    })
                $exitCode = 0
                break
            }

            # =====================================================================================
            # Start: submit a plan.
            # =====================================================================================
            'Start' {
                $total = 6
                Write-EmmStep 1 $total 'Loading the plan' -Icon Plan
                $wanted = if ($Workload) { $Workload } else { $null }
                if (-not $PlanPath) {
                    $PlanPath = Find-EmmPlan -OutputPath $settings.Report.OutputPath -Workload $(if ($wanted) { $wanted } else { 'All' })
                    if (-not $PlanPath -and -not $wanted) {
                        foreach ($w in 'System', 'User') { if (-not $PlanPath) { $PlanPath = Find-EmmPlan -OutputPath $settings.Report.OutputPath -Workload $w } }
                    }
                    if (-not $PlanPath) { throw "No plan found in $($settings.Report.OutputPath)$(if ($wanted) { " for the workload $wanted" }). Run -Mode Plan first." }
                }
                $plan = Import-EmmPlan -Path $PlanPath
                $planWorkload = [string]$plan.Meta.Workload
                if (-not $wanted) { $wanted = $planWorkload }
                if ($planWorkload -ne 'All' -and $planWorkload -ne $wanted) { throw "The plan $($plan.Meta.PlanId) covers the workload $planWorkload, not $wanted." }
                $created = [datetime]::ParseExact([string]$plan.Meta.Created, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
                $age = (Get-Date) - $created
                Write-EmmItem Ok ('Plan {0} {1} {2} {1} created {3} ({4} ago)' -f $plan.Meta.PlanId, $dot, $plan.Meta.Selection, $created.ToString('yyyy-MM-dd HH:mm'), (Format-EmmDuration $age.TotalSeconds)) -Icon Plan
                Write-EmmItem Info ('{0}' -f $plan.Directory) -Icon Folder
                if ($age.TotalDays -gt $settings.Plan.MaxPlanAgeDays) {
                    if (-not $Force) { throw ("The plan is {0:0} day(s) old (Plan.MaxPlanAgeDays = {1}): create a new plan, or use -Force." -f $age.TotalDays, $settings.Plan.MaxPlanAgeDays) }
                    Write-EmmItem Warn ("The plan is {0:0} day(s) old: accepted with -Force." -f $age.TotalDays)
                }
                $selectedRows = @($plan.Rows | Where-Object { $wanted -eq 'All' -or $_.Workload -eq $wanted })
                Write-EmmItem Info ('{0} mailbox(es) of the plan in the workload {1}: {2} system, {3} user' -f $selectedRows.Count, $wanted, @($selectedRows | Where-Object { $_.Workload -eq 'System' }).Count, @($selectedRows | Where-Object { $_.Workload -eq 'User' }).Count) -Icon Mailbox
                if (-not $selectedRows.Count) { throw "The plan has no mailbox in the workload $wanted." }

                $connection = Connect-Step 2 $total

                Write-EmmStep 3 $total 'Checking the plan against the organisation' -Icon Search
                $pre = Get-EmmStartPreflight -Settings $settings -Plan $plan -Workload $wanted
                $items = @($pre.Items)
                $byBatch = @($items | Group-Object { $_.Row.Batch } | Sort-Object { if ($_.Group[0].Row.Workload -eq 'System') { '0' } else { '1' + $_.Name } })
                Write-EmmTable -Columns ([ordered]@{ Batch = -22; Planned = 8; Submit = 7; Skip = 6; Volume = 10; Decision = -30 }) -Rows @($byBatch | ForEach-Object {
                        $g = @($_.Group); $name = $_.Name; $bd = @($pre.Batches | Where-Object { $_.Name -eq $name })
                        $decision = if ($g[0].Row.Workload -eq 'System') { 'individual move requests' } elseif ($bd.Count) { $bd[0].Action + $(if ($bd[0].Existing) { " ($($bd[0].Existing))" } else { '' }) } else { '' }
                        $submitCount = @($g | Where-Object { $_.Decision -eq 'Submit' }).Count
                        [pscustomobject]@{ Status = $(if ($submitCount -eq $g.Count) { 'Ok' } elseif ($submitCount) { 'Warn' } else { 'Fail' }); Batch = $name; Planned = $g.Count; Submit = $submitCount; Skip = $g.Count - $submitCount
                            Volume = Format-EmmSize ([double](@($g | Where-Object { $_.Decision -eq 'Submit' } | ForEach-Object { $_.Row.MoveSizeMB } | Measure-Object -Sum)[0].Sum)); Decision = $decision }
                    })
                foreach ($r in @($items | Where-Object { $_.Decision -eq 'Skip' } | Group-Object Reason | Sort-Object Count -Descending | Select-Object -First 8)) {
                    $st = $r.Group[0].Status
                    Write-EmmItem $(if ($st -eq 'Failed') { 'Fail' } elseif ($st -eq 'AlreadyDone') { 'Ok' } else { 'Skip' }) ('{0,5} x {1}' -f $r.Count, $r.Name)
                }
                $reduced = @($items | Where-Object { $_.Decision -eq 'Submit' -and $_.Reason -like 'Move type reduced*' }).Count
                if ($reduced) { Write-EmmItem Info ('{0} mailbox(es): move type reduced (one part already moved)' -f $reduced) }
                foreach ($pb in @($pre.PreviousBatches)) { Write-EmmItem Info ('Earlier batch {0} removed first ({1}): {2} planned mailbox(es) are still its migration users' -f $pb.Name, $pb.Reason, $pb.Mailboxes) -Icon Broom }
                $orphanUsers = @($items | Where-Object { $_.Decision -eq 'Submit' -and $_.RemoveMigrationUser }).Count
                if ($orphanUsers) { Write-EmmItem Info ('{0} orphan migration user(s) of earlier batches removed first' -f $orphanUsers) -Icon Broom }
                foreach ($w in $pre.Warnings) { Write-EmmItem Warn $w }

                $submitItems = @($items | Where-Object { $_.Decision -eq 'Submit' })
                $submitVolume = [double](@($submitItems | ForEach-Object { $_.Row.MoveSizeMB } | Measure-Object -Sum)[0].Sum)
                $sysCount = @($submitItems | Where-Object { $_.Row.Workload -eq 'System' }).Count
                $userCount = $submitItems.Count - $sysCount

                Write-EmmStep 4 $total 'Confirmation' -Icon Lock
                if (-not $submitItems.Count) {
                    Write-EmmItem Warn 'Nothing to submit.'
                    $go = $false
                } else {
                    $question = 'Submit {0} mailbox move(s) ({1}): {2} system move request(s) and {3} user mailbox(es) in {4} migration batch(es){5}?' -f $submitItems.Count, (Format-EmmSize $submitVolume), $sysCount, $userCount, @($pre.Batches | Where-Object { $_.Action -ne 'Skip' }).Count, $(if (@($pre.PreviousBatches).Count) { ', after removing {0} finished earlier batch(es)' -f @($pre.PreviousBatches).Count } else { '' })
                    Write-EmmItem Info $question -Icon Next
                    $go = Confirm-Change $question
                    if (-not $go) {
                        Write-EmmSummary -Title 'Cancelled: nothing submitted' -Status Warn -Values ([ordered]@{ 'Plan' = @('Plan', $plan.Meta.PlanId); 'Duration' = @('Clock', (Get-RunDuration)) })
                        $exitCode = 3
                        break
                    }
                }

                Write-EmmStep 5 $total $(if ($simulate) { 'Submitting (simulation)' } else { 'Submitting the moves' }) -Icon Rocket
                if ($go) { Start-EmmMigration -Settings $settings -Preflight $pre -Simulate:$simulate }
                else { Start-EmmMigration -Settings $settings -Preflight ([pscustomobject]@{ Items = $items; Batches = @() }) -Simulate:$simulate }

                Write-EmmStep 6 $total 'Writing the report' -Icon Report
                $actions = @(Get-EmmActions)
                $batchRows = @($byBatch | ForEach-Object {
                        $g = @($_.Group); $name = $_.Name; $bd = @($pre.Batches | Where-Object { $_.Name -eq $name })
                        [pscustomobject][ordered]@{
                            Name = $name; Workload = $g[0].Row.Workload; Action = $(if ($g[0].Row.Workload -eq 'System') { 'Submit' } elseif ($bd.Count) { $bd[0].Action } else { '' })
                            Reason = $(if ($bd.Count) { $bd[0].Reason } else { '' }); Mailboxes = $g.Count; Submitted = @($g | Where-Object { $_.Decision -eq 'Submit' }).Count
                            SizeMB = [Math]::Round([double](@($g | ForEach-Object { $_.Row.MoveSizeMB } | Measure-Object -Sum)[0].Sum), 2)
                        }
                    })
                $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode $(if ($simulate) { 'Start-WhatIf' } else { 'Start' }) -Suffix $wanted
                $n = Get-ActionCount
                $meta = [ordered]@{ workload = $wanted; selection = $plan.Meta.Selection; planId = $plan.Meta.PlanId; simulation = $simulate; connection = $connection.Method; account = $connection.Account
                    warnings = @($pre.Warnings); duration = (Get-RunDuration)
                    next = $(if ($simulate) { "$tool -Mode Start -Workload $wanted" } else { "$tool -Mode Status -Follow" })
                    nextText = $(if ($simulate) { 'Simulation finished: run the same command without -WhatIf to submit the moves.' } else { 'Follow the synchronisation; user batches stop when synchronised and wait for -Mode Complete.' }) }
                $files = New-EmmReport -Kind Start -Settings $settings -Directory $dir -Meta $meta -Tables ([ordered]@{ actions = $actions; batches = $batchRows }) -Csv ([ordered]@{ 'Actions.csv' = 'actions' })
                Write-Files $files
                $failedOrSkipped = $n.Failed + @($items | Where-Object { $_.Decision -eq 'Skip' -and $_.Status -eq 'Skipped' }).Count
                Write-EmmSummary -Title $(if ($simulate) { 'Simulation finished' } elseif ($n.Failed) { 'Started, with errors' } else { 'Migration started' }) -Status $(if ($n.Failed) { 'Fail' } elseif ($failedOrSkipped) { 'Warn' } else { 'Ok' }) -Values ([ordered]@{
                        'Plan'      = @('Plan', ('{0} {1} workload {2}' -f $plan.Meta.PlanId, $dot, $wanted))
                        'Submitted' = @($(if ($simulate) { 'Simulate' } else { 'Rocket' }), ('{0} mailbox(es) {1} {2} ({3} system, {4} user)' -f $(if ($go) { $submitItems.Count } else { 0 }), $dot, (Format-EmmSize $(if ($go) { $submitVolume } else { 0 })), $(if ($go) { $sysCount } else { 0 }), $(if ($go) { $userCount } else { 0 })))
                        'Not sent'  = @('Skip', ('{0} skipped {1} {2} already done {1} {3} failed' -f @($items | Where-Object { $_.Status -eq 'Skipped' }).Count, $dot, @($items | Where-Object { $_.Status -eq 'AlreadyDone' }).Count, $n.Failed))
                        'Report'    = @('Folder', $dir)
                        'Next'      = @('Next', $meta.next)
                        'Duration'  = @('Clock', (Get-RunDuration))
                    })
                $exitCode = if ($n.Failed -or $failedOrSkipped) { 2 } else { 0 }
                break
            }

            # =====================================================================================
            # Status: real state, optionally followed until the end.
            # =====================================================================================
            'Status' {
                $total = 3
                $connection = Connect-Step 1 $total
                Write-EmmStep 2 $total 'Reading the migration state' -Icon Sync
                $objects = Get-EmmMigrationObjects -Settings $settings
                $existing = @(Get-EmmManagedName -Objects $objects)
                $filter = if ($Batch) { Select-EmmBatchName -Settings $settings -Batch $Batch -Existing $existing } else { $existing }
                if ($Workload -eq 'System') { $filter = @($filter | Where-Object { $_ -eq $settings.System.BatchName }) }
                elseif ($Workload -eq 'User') { $filter = @($filter | Where-Object { $_ -ne $settings.System.BatchName }) }
                $missing = @($filter | Where-Object { $_ -notin $existing })
                foreach ($m in $missing) { Write-EmmItem Warn "$m`: no batch or move request of this tool with this name." }
                $filter = @($filter | Where-Object { $_ -in $existing })
                if (-not $filter.Count) {
                    Write-EmmItem Warn 'No migration batch or move request of this tool to show.'
                    Write-EmmSummary -Title 'Nothing to show' -Status Warn -Values ([ordered]@{ 'Next' = @('Next', "$tool -Mode Plan"); 'Duration' = @('Clock', (Get-RunDuration)) })
                    $exitCode = 0
                    break
                }
                $label = if ($Batch -or $Workload) { $filter -join ', ' } else { '' }
                $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode 'Status'
                Write-EmmStep 3 $total $(if ($Follow) { 'Report and follow-up' } else { 'Writing the report' }) -Icon Report
                if ($Follow) { $round = Invoke-Follow -Filter $filter -Directory $dir -FilterLabel $label }
                else { $round = Invoke-StatusRound -Filter $filter -Directory $dir -RefreshSeconds 0 -FilterLabel $label }
                Write-StatusTable $round.Status.Batches
                Write-Files $round.Files
                $rows = @($round.Status.Rows)
                $failed = @($rows | Where-Object { $_.Group -eq 'Failed' }).Count
                Write-EmmSummary -Title 'Migration status' -Status $(if ($failed) { 'Warn' } else { 'Ok' }) -Values ([ordered]@{
                        'Mailboxes' = @('Mailbox', ('{0} {1} {2} completed {1} {3} ready {1} {4} in progress {1} {5} failed' -f $rows.Count, $dot, @($rows | Where-Object { $_.Group -eq 'Completed' }).Count, @($rows | Where-Object { $_.Group -eq 'Ready' }).Count, @($rows | Where-Object { $_.Group -eq 'InProgress' }).Count, $failed))
                        'Report'    = @('Folder', $dir)
                        'Next'      = @('Next', $(if ($round.Next[0]) { $round.Next[0] } else { '-' }))
                        'Duration'  = @('Clock', (Get-RunDuration))
                    })
                $exitCode = if ($failed) { 2 } else { 0 }
                break
            }

            # =====================================================================================
            # Complete: completion now or scheduled, optionally followed.
            # =====================================================================================
            'Complete' {
                $total = 5
                $connection = Connect-Step 1 $total
                Write-EmmStep 2 $total 'Selecting the batches' -Icon Batch
                $objects = Get-EmmMigrationObjects -Settings $settings
                # 'All' = every batch and public folder label of this tool; the system moves complete by themselves (-Batch System if needed).
                $names = @(Select-EmmBatchName -Settings $settings -Batch $Batch -Existing @(Get-EmmManagedName -Objects $objects | Where-Object { $_ -ne $settings.System.BatchName }))
                $rows = foreach ($n in $names) {
                    $b = @($objects.ToolBatches | Where-Object { $_.Name -eq $n })
                    $labelMoves = @($objects.ToolMoves | Where-Object { $_.ToolBatch -eq $n })
                    $status = if ($b.Count) { $b[0].Status } elseif ($labelMoves.Count) { 'move requests' } else { 'not found' }
                    $ok = (-not $b.Count -and $labelMoves.Count) -or ($b.Count -and $b[0].Status -eq 'Synced') -or ($PSBoundParameters.ContainsKey('CompleteAfter') -and $b.Count -and $b[0].Status -eq 'Syncing')
                    [pscustomobject]@{ Status = $(if ($ok) { 'Ok' } elseif ($b.Count -and $b[0].Status -in 'Completed', 'CompletedWithErrors') { 'Skip' } else { 'Warn' }); Batch = $n; State = $status; Mailboxes = $(if ($b.Count) { $b[0].Total } else { $labelMoves.Count }); Synced = $(if ($b.Count) { $b[0].Synced } else { @($labelMoves | Where-Object { $_.Status -in 'AutoSuspended', 'Synced' }).Count }) }
                }
                Write-EmmTable -Columns ([ordered]@{ Batch = -22; State = -22; Mailboxes = 9; Synced = 7 }) -Rows @($rows)
                if (-not $names.Count) { throw 'No batch selected.' }
                $when = if ($PSBoundParameters.ContainsKey('CompleteAfter')) { $CompleteAfter } else { $null }

                Write-EmmStep 3 $total 'Confirmation' -Icon Lock
                $question = if ($when) { 'Schedule the completion of {0} at {1}?' -f ($names -join ', '), $when.ToString('yyyy-MM-dd HH:mm') } else { 'Complete {0} now (users switch to the target databases)?' -f ($names -join ', ') }
                Write-EmmItem Info $question -Icon Next
                if (-not (Confirm-Change $question)) {
                    Write-EmmSummary -Title 'Cancelled: nothing completed' -Status Warn -Values ([ordered]@{ 'Batches' = @('Batch', ($names -join ', ')); 'Duration' = @('Clock', (Get-RunDuration)) })
                    $exitCode = 3
                    break
                }

                Write-EmmStep 4 $total $(if ($when) { 'Scheduling the completion' } else { 'Completing' }) -Icon Flag
                Complete-EmmBatch -Settings $settings -BatchName $names -CompleteAfter $when -Simulate:$simulate

                Write-EmmStep 5 $total 'Writing the report' -Icon Report
                $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode $(if ($simulate) { 'Complete-WhatIf' } else { 'Complete' })
                $n = Get-ActionCount
                $statusCommand = "$tool -Mode Status -Batch $(($names | ForEach-Object { $_ -replace ('^' + [regex]::Escape($settings.Plan.BatchNamePrefix)), '' }) -join ',') -Follow"
                $meta = [ordered]@{ simulation = $simulate; batchFilter = ($names -join ', '); connection = $connection.Method; account = $connection.Account; duration = (Get-RunDuration)
                    next = $statusCommand; nextText = $(if ($when) { "Completion scheduled at $($when.ToString('yyyy-MM-dd HH:mm')): follow it." } else { 'Follow the completion until every mailbox is completed.' }) }
                $files = New-EmmReport -Kind Complete -Settings $settings -Directory $dir -Meta $meta -Tables ([ordered]@{ actions = @(Get-EmmActions) }) -Csv ([ordered]@{ 'Actions.csv' = 'actions' })
                Write-Files $files
                if ($Follow -and -not $simulate) {
                    Write-EmmItem Info 'Follow-up of the completed batches:' -Icon Sync
                    $round = Invoke-Follow -Filter $names -Directory $dir -FilterLabel ($names -join ', ')
                    Write-StatusTable $round.Status.Batches
                }
                Write-EmmSummary -Title $(if ($simulate) { 'Simulation finished' } elseif ($n.Failed) { 'Completion with errors' } elseif ($when) { 'Completion scheduled' } else { 'Completion requested' }) -Status $(if ($n.Failed) { 'Fail' } elseif ($n.Skipped) { 'Warn' } else { 'Ok' }) -Values ([ordered]@{
                        'Batches'  = @('Batch', ('{0} {1} {2} done {1} {3} skipped {1} {4} failed' -f ($names -join ', '), $dot, ($n.Success + $n.Simulated + $n.AlreadyDone), $n.Skipped, $n.Failed))
                        'Report'   = @('Folder', $dir)
                        'Next'     = @('Next', $statusCommand)
                        'Duration' = @('Clock', (Get-RunDuration))
                    })
                $exitCode = if ($n.Failed -or $n.Skipped) { 2 } else { 0 }
                break
            }

            # =====================================================================================
            # Cleanup: remove what this tool created and no longer moves anything.
            # =====================================================================================
            'Cleanup' {
                $total = 5
                $connection = Connect-Step 1 $total
                Write-EmmStep 2 $total 'Finding the finished objects of this tool' -Icon Search
                $objects = Get-EmmMigrationObjects -Settings $settings
                $managed = @(Get-EmmManagedName -Objects $objects)
                $names = if ($Batch) { @(Select-EmmBatchName -Settings $settings -Batch $Batch -Existing $managed) } else { $null }
                if ($Workload -eq 'System') { $names = @($settings.System.BatchName) }
                elseif ($Workload -eq 'User') { $names = @($(if ($names) { $names } else { $managed }) | Where-Object { $_ -ne $settings.System.BatchName }); if (-not $names.Count) { $names = @('-none-') } }
                $cleanup = Get-EmmCleanupPlan -Settings $settings -BatchName $names
                foreach ($b in $cleanup.Batches) { Write-EmmItem Ok ('{0,-22} remove {1} {2}' -f $b.Batch.Name, $dot, $b.Reason) -Icon Broom }
                foreach ($b in $cleanup.KeptBatches) { Write-EmmItem Skip ('{0,-22} kept   {1} {2}' -f $b.Batch.Name, $dot, $b.Reason) }
                Write-EmmItem Info ('{0} finished move request(s) to remove {1} {2} not finished, kept' -f $cleanup.MoveRequests.Count, $dot, $cleanup.KeptMoves) -Icon Move
                if ($settings.Cleanup.IncludeFailed) { Write-EmmItem Warn 'Cleanup.IncludeFailed: failed batches and move requests are removed too.' }

                Write-EmmStep 3 $total 'Confirmation' -Icon Lock
                if (-not $cleanup.Batches.Count -and -not $cleanup.MoveRequests.Count -and -not $settings.Cleanup.RemoveOrphanMigrationUsers) {
                    Write-EmmItem Ok 'Nothing to remove.'
                    $go = $false
                } else {
                    $question = 'Remove {0} migration batch(es), {1} move request(s){2}?' -f $cleanup.Batches.Count, $cleanup.MoveRequests.Count, $(if ($settings.Cleanup.RemoveOrphanMigrationUsers) { ' and the orphan migration users' } else { '' })
                    Write-EmmItem Info $question -Icon Next
                    $go = Confirm-Change $question
                    if (-not $go) {
                        Write-EmmSummary -Title 'Cancelled: nothing removed' -Status Warn -Values ([ordered]@{ 'Duration' = @('Clock', (Get-RunDuration)) })
                        $exitCode = 3
                        break
                    }
                }
                Write-EmmStep 4 $total 'Removing' -Icon Broom
                $orphans = 0
                if ($go) { $orphans = Invoke-EmmCleanup -Settings $settings -CleanupPlan $cleanup -Simulate:$simulate } else { Write-EmmItem Skip 'Nothing to do.' }

                Write-EmmStep 5 $total 'Writing the report' -Icon Report
                $dir = New-EmmRunDirectory -OutputPath $settings.Report.OutputPath -Mode $(if ($simulate) { 'Cleanup-WhatIf' } else { 'Cleanup' })
                $n = Get-ActionCount
                $meta = [ordered]@{ simulation = $simulate; connection = $connection.Method; account = $connection.Account; batchFilter = $(if ($names) { $names -join ', ' } else { '' }); duration = (Get-RunDuration) }
                $files = New-EmmReport -Kind Cleanup -Settings $settings -Directory $dir -Meta $meta -Tables ([ordered]@{ actions = @(Get-EmmActions) }) -Csv ([ordered]@{ 'Actions.csv' = 'actions' })
                Write-Files $files
                Write-EmmSummary -Title $(if ($simulate) { 'Simulation finished' } elseif ($n.Failed) { 'Cleanup with errors' } else { 'Cleanup finished' }) -Status $(if ($n.Failed) { 'Fail' } else { 'Ok' }) -Values ([ordered]@{
                        'Removed'  = @('Broom', ('{0} batch(es) {1} {2} move request(s) {1} {3} orphan migration user(s)' -f $(if ($go) { $cleanup.Batches.Count } else { 0 }), $dot, $(if ($go) { $cleanup.MoveRequests.Count } else { 0 }), $orphans))
                        'Kept'     = @('Lock', ('{0} batch(es) {1} {2} move request(s) not finished' -f $cleanup.KeptBatches.Count, $dot, $cleanup.KeptMoves))
                        'Report'   = @('Folder', $dir)
                        'Duration' = @('Clock', (Get-RunDuration))
                    })
                $exitCode = if ($n.Failed) { 2 } else { 0 }
                break
            }
        }
    } while ($false)
}
catch {
    $message = $_.Exception.Message
    Write-Host ''
    if (Get-Command Write-EmmSummary -ErrorAction SilentlyContinue) {
        Write-EmmSummary -Title 'Execution stopped' -Values ([ordered]@{ Error = @('Fail', $message); Duration = @('Clock', (Format-EmmDuration $clock.Elapsed.TotalSeconds)) }) -Status Fail
    } else {
        Write-Host "  [ERROR] $message"
    }
    if (Get-Command Write-EmmLog -ErrorAction SilentlyContinue) { Write-EmmLog -Level 'ERROR' -Message ($message + "`n" + $_.ScriptStackTrace) }
    $exitCode = 1
}
finally {
    if (Get-Command Disconnect-EmmExchange -ErrorAction SilentlyContinue) { Disconnect-EmmExchange }
    if (Get-Command Write-EmmLog -ErrorAction SilentlyContinue) {
        Write-EmmLog -Level 'INFO' -Message ('Exit code {0} after {1:0.0} s' -f $exitCode, $clock.Elapsed.TotalSeconds)
        Stop-EmmLog
    }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
