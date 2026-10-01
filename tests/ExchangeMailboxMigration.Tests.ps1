#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Exchange Mailbox Migration - automated tests (Pester 5 or later).
    Author  : Nicolas Fabert
    Version : 2.0.0

    Run:  Invoke-Pester -Path .\tests\ExchangeMailboxMigration.Tests.ps1 -Output Detailed

    No Exchange server is needed: tests\FakeExchange.ps1 provides the Exchange cmdlets over a
    fictitious in-memory organisation (contoso.com). The whole lifecycle with the real entry script,
    in Windows PowerShell 5.1 too, is covered by tests\Invoke-EndToEnd.ps1.
#>

BeforeAll {
    $script:Root = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:Root 'ExchangeMailboxMigration.psd1') -Force
    . (Join-Path $PSScriptRoot 'FakeExchange.ps1')
    $script:Temp = Join-Path ([IO.Path]::GetTempPath()) ('EmmTests-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($script:Temp)

    function New-TestConfig {
        <# The delivered configuration file with replacements (regex -> replacement), written in the temp folder. #>
        param([System.Collections.IDictionary]$Replace = @{})
        $text = [IO.File]::ReadAllText((Join-Path $script:Root 'config\ExchangeMailboxMigration.config.psd1'))
        $base = [ordered]@{
            "(?m)^(\s*OutputPath\s*=\s*)'[^']*'"            = "`${1}'$((Join-Path $script:Temp 'reports').Replace("'", "''"))'"
            "(?m)^(\s*Path\s*=\s*)'\.\\logs'"               = "`${1}'$((Join-Path $script:Temp 'logs').Replace("'", "''"))'"
            "(?m)^(\s*TargetDatabasePattern\s*=\s*)'[^']*'" = "`${1}'^DB-0[1-4]`$`$'"
            "(?m)^(\s*PollSeconds\s*=\s*)\d+"               = '${1}1'
            "(?m)^(\s*SettleSeconds\s*=\s*)\d+"             = '${1}0'
        }
        foreach ($k in $Replace.Keys) { $base[$k] = $Replace[$k] }
        foreach ($k in $base.Keys) {
            if ([regex]::Matches($text, $k).Count -ne 1) { throw "Test configuration: '$k' must match once." }
            $text = [regex]::Replace($text, $k, $base[$k])
        }
        $path = Join-Path $script:Temp ('config-' + [guid]::NewGuid().ToString('N') + '.psd1')
        [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($true)))
        return $path
    }
    function New-TestSettings([System.Collections.IDictionary]$Replace = @{}) { Import-EmmConfiguration -Path (New-TestConfig $Replace) -Root $script:Root }
    function Get-TestInventory($Settings, $Scope) {
        $db = Get-EmmDatabaseInventory -Settings $Settings
        $mbx = Get-EmmMailboxInventory -Settings $Settings -DatabaseInventory $db -Scope $Scope -Quiet 6>$null
        return [pscustomobject]@{ Databases = $db; Mailboxes = $mbx.Mailboxes }
    }
    function Mbx($Inventory, [string]$Alias) { @($Inventory.Mailboxes | Where-Object { $_.Alias -eq $Alias })[0] }
}

AfterAll {
    if ($script:Temp -and (Test-Path -LiteralPath $script:Temp)) { Remove-Item -LiteralPath $script:Temp -Recurse -Force }
}

Describe 'Configuration' {
    It 'accepts the delivered configuration file' {
        $s = Import-EmmConfiguration -Path (Join-Path $script:Root 'config\ExchangeMailboxMigration.config.psd1') -Root $script:Root
        $s.Scope.Workload | Should -Be 'All'
        $s.Scope.SystemMailboxTypes | Should -Contain 'ArbitrationMailbox'
        $s.Scope.UserMailboxTypes | Should -Contain 'Archive'
        [IO.Path]::IsPathRooted($s.Report.OutputPath) | Should -BeTrue
    }
    It 'reports every problem at once' {
        $path = New-TestConfig ([ordered]@{
                "(?m)^(\s*BadItemLimit\s*=\s*)\d+"          = '${1}60'
                "(?m)^(\s*SourceDatabasePattern\s*=\s*)'[^']*'" = "`${1}'^(DB'"
                "(?m)^(\s*BatchCount\s*=\s*)\d+"            = '${1}120'
            })
        { Import-EmmConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*AcceptLargeDataLoss*'
        try { Import-EmmConfiguration -Path $path -Root $script:Root } catch { $m = $_.Exception.Message }
        $m | Should -BeLike '*SourceDatabasePattern is not a valid regular expression*'
        $m | Should -BeLike '*Plan.BatchCount must be a whole number between 1 and 99*'
    }
    It 'refuses monitoring mailboxes in the scope' {
        $path = New-TestConfig ([ordered]@{ "(?m)^(\s*SystemMailboxTypes\s*=\s*)@\([^)]*\)" = "`${1}@('ArbitrationMailbox', 'MonitoringMailbox')" })
        { Import-EmmConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*never moved*'
    }
}

Describe 'Scope (workload and mailbox types)' {
    BeforeAll { $script:S = New-TestSettings }
    It 'All: system types and user types with archives' {
        $sc = Resolve-EmmScope -Settings $script:S
        $sc.Workload | Should -Be 'All'
        $sc.SystemTypes | Should -Contain 'ArbitrationMailbox'
        $sc.UserPrimaryTypes | Should -Contain 'SharedMailbox'
        $sc.UserPrimaryTypes | Should -Not -Contain 'Archive'
        $sc.IncludeArchives | Should -BeTrue
    }
    It 'System: no user type' {
        $sc = Resolve-EmmScope -Settings $script:S -Workload System
        @($sc.UserPrimaryTypes).Count | Should -Be 0
        $sc.IncludeArchives | Should -BeFalse
    }
    It 'deduces the workload from -MailboxType' {
        (Resolve-EmmScope -Settings $script:S -MailboxType 'SharedMailbox', 'RoomMailbox').Workload | Should -Be 'User'
        (Resolve-EmmScope -Settings $script:S -MailboxType 'ArbitrationMailbox').Workload | Should -Be 'System'
        (Resolve-EmmScope -Settings $script:S -MailboxType 'DiscoveryMailbox,UserMailbox').Workload | Should -Be 'All'
    }
    It 'Archive alone selects the archives of every user type' {
        $sc = Resolve-EmmScope -Settings $script:S -MailboxType Archive
        $sc.IncludeArchives | Should -BeTrue
        @($sc.UserPrimaryTypes).Count | Should -Be 0
        $sc.ArchiveOwnerTypes | Should -Contain 'UserMailbox'
        $sc.ArchiveOwnerTypes | Should -Not -Contain 'PublicFolderMailbox'
    }
    It 'never selects monitoring mailboxes' {
        { Resolve-EmmScope -Settings $script:S -MailboxType MonitoringMailbox } | Should -Throw -ExpectedMessage '*never moved*'
        { Resolve-EmmScope -Settings $script:S -MailboxType Nonsense } | Should -Throw -ExpectedMessage '*Unknown mailbox type*'
    }
}

Describe 'Database classification' {
    It 'never treats a target database as a source' {
        $s = New-TestSettings ([ordered]@{ "(?m)^(\s*SourceDatabasePattern\s*=\s*)'[^']*'" = "`${1}'^DB'" })
        $c = Get-EmmDatabaseClassification -Name 'DB01', 'DB-01', 'Other' -Settings $s
        $c.Roles['DB01'].Role | Should -Be 'Source'
        $c.Roles['DB-01'].Role | Should -Be 'Target'
        $c.Roles['Other'].Role | Should -Be 'Other'
        $c.Warnings -join ' ' | Should -BeLike '*DB-01 matches both*'
    }
    It 'applies DatabaseMap and reports a missing target' {
        $s = New-TestSettings ([ordered]@{ "(?m)^(\s*#\s*'DB01'\s*=\s*'DB-01')" = "            'DB01' = 'DB-09'" })
        $c = Get-EmmDatabaseClassification -Name 'DB01', 'DB-01' -Settings $s
        $c.Roles['DB01'].MapTarget | Should -Be 'DB-09'
        $c.Warnings -join ' ' | Should -BeLike '*DB-09 does not exist*'
    }
}

Describe 'Sizes and JSON' {
    It 'reads Exchange sizes whatever the regional settings' {
        ConvertTo-EmmMB '1.5 GB (1,610,612,736 bytes)' | Should -Be 1536
        ConvertTo-EmmMB ("1,5 Go (1{0}610{0}612{0}736 octets)" -f [char]0x00A0) | Should -Be 1536
        ConvertTo-EmmMB $null | Should -Be 0
        ConvertTo-EmmMB ([long]1048576) | Should -Be 1
    }
    It 'writes JSON that is safe inside a script element and culture independent' {
        $previous = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            $json = ConvertTo-EmmJson ([ordered]@{ a = '</script><b>&'; n = 1.5; list = @(1, 'x', $true, $null) })
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $previous }
        $json | Should -Not -BeLike '*</script>*'
        $json | Should -BeLike '*"n":1.5*'
        $json | Should -BeLike '*[1,"x",true,null]*'
    }
}

Describe 'Inventory and selection (fictitious organisation)' {
    BeforeAll {
        New-EmmFakeOrg
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*ExcludeMailboxes\s*=\s*)@\(\)" = "`${1}@('ceo@contoso.com')" })
        $script:I = Get-TestInventory $script:S (Resolve-EmmScope -Settings $script:S)
    }
    It 'excludes every monitoring mailbox' {
        $mon = @($script:I.Mailboxes | Where-Object { $_.Category -eq 'Monitoring' })
        $mon.Count | Should -BeGreaterThan 0
        @($mon | Where-Object { $_.Status -ne 'Excluded' }).Count | Should -Be 0
    }
    It 'recognises a HealthMailbox even with another recipient type' {
        $r = ConvertTo-EmmMailboxRecord -Mailbox ([pscustomobject]@{ Name = 'HealthMailbox0123'; RecipientTypeDetails = 'UserMailbox'; Database = 'DB01' })
        $r.Category | Should -Be 'Monitoring'
    }
    It 'decides the move type of each mailbox' {
        (Mbx $script:I 'archive.left').MoveType | Should -Be 'ArchiveOnly'
        (Mbx $script:I 'archive.moved').MoveType | Should -Be 'PrimaryOnly'
        (Mbx $script:I 'already.moved').Status | Should -Be 'OnTarget'
        (Mbx $script:I 'other.db').Status | Should -Be 'Outside'
        (Mbx $script:I 'ceo').Status | Should -Be 'Excluded'
        @($script:I.Mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'TeamMailbox' })[0].Status | Should -Be 'NotSelected'
    }
    It 'measures the moved size (items + recoverable items)' {
        $m = @($script:I.Mailboxes | Where-Object { $_.Name -eq 'Accounting' })[0]
        $m.MoveSizeMB | Should -BeGreaterThan 0
        [Math]::Abs($m.MoveSizeMB - ($m.PrimarySizeMB + $m.ArchiveSizeMB)) | Should -BeLessThan 0.05
    }
    It 'keeps user mailboxes out of the System workload' {
        $sys = Get-TestInventory $script:S (Resolve-EmmScope -Settings $script:S -Workload System)
        @($sys.Mailboxes | Where-Object { $_.Status -eq 'ToMove' -and $_.Category -ne 'System' }).Count | Should -Be 0
        @($sys.Mailboxes | Where-Object { $_.Status -eq 'ToMove' -and $_.RecipientTypeDetails -eq 'ArbitrationMailbox' }).Count | Should -Be 5
    }
}

Describe 'Plan' {
    BeforeAll {
        New-EmmFakeOrg
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}5' })
        $script:Scope = Resolve-EmmScope -Settings $script:S
        $script:I = Get-TestInventory $script:S $script:Scope
        $script:P = New-EmmPlan -Settings $script:S -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
    }
    It 'balances the user batches by count and volume' {
        $user = @($script:P.Batches | Where-Object { $_.Workload -eq 'User' })
        $user.Count | Should -Be 5
        $counts = @($user | ForEach-Object { $_.Mailboxes })
        (($counts | Measure-Object -Maximum).Maximum - ($counts | Measure-Object -Minimum).Minimum) | Should -BeLessOrEqual 1
        $sizes = @($user | ForEach-Object { $_.SizeMB })
        (($sizes | Measure-Object -Maximum).Maximum / ($sizes | Measure-Object -Minimum).Minimum) | Should -BeLessThan 1.1
    }
    It 'keeps system mailboxes out of the batches' {
        @($script:P.Rows | Where-Object { $_.Workload -eq 'System' -and $_.Batch -ne $script:S.System.BatchName }).Count | Should -Be 0
        @($script:P.Rows | Where-Object { $_.Workload -eq 'User' -and $_.Batch -notmatch '^Batch\d\d$' }).Count | Should -Be 0
    }
    It 'targets only target databases and keeps archives with their primary' {
        @($script:P.Rows | Where-Object { ($_.TargetDatabase -and $_.TargetDatabase -notmatch '^DB-0[1-4]$') -or ($_.TargetArchiveDatabase -and $_.TargetArchiveDatabase -notmatch '^DB-0[1-4]$') }).Count | Should -Be 0
        @($script:P.Rows | Where-Object { $_.MoveType -eq 'PrimaryAndArchive' -and $_.TargetArchiveDatabase -ne $_.TargetDatabase }).Count | Should -Be 0
        (@($script:P.Rows | Where-Object { $_.Alias -eq 'archive.left' })[0]).TargetArchiveDatabase | Should -Be 'DB-01'
    }
    It 'never plans a monitoring mailbox' {
        @($script:P.Rows | Where-Object { $_.Name -like 'HealthMailbox*' }).Count | Should -Be 0
    }
    It 'sends archives to ArchiveTargetDatabasePattern when set' {
        $s = New-TestSettings ([ordered]@{ "(?m)^(\s*ArchiveTargetDatabasePattern\s*=\s*)''" = "`${1}'^DB-04`$`$'" })
        $p = New-EmmPlan -Settings $s -DatabaseInventory (Get-EmmDatabaseInventory -Settings $s) -Mailbox $script:I.Mailboxes -Scope $script:Scope
        @($p.Rows | Where-Object { $_.TargetArchiveDatabase -and $_.TargetArchiveDatabase -ne 'DB-04' }).Count | Should -Be 0
    }
    It 'PerSourceDatabase: one batch per source database' {
        $s = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchStrategy\s*=\s*)'[^']*'" = "`${1}'PerSourceDatabase'" })
        $p = New-EmmPlan -Settings $s -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
        $key = { param($r) if ($r.MoveType -eq 'ArchiveOnly') { $r.SourceArchiveDatabase } else { $r.SourceDatabase } }
        foreach ($g in @($p.Rows | Where-Object { $_.Workload -eq 'User' } | Group-Object Batch)) { @($g.Group | ForEach-Object { & $key $_ } | Sort-Object -Unique).Count | Should -Be 1 }
        @($p.Batches | Where-Object { $_.Workload -eq 'User' }).Count | Should -Be 4
    }
    It 'saves and reads back the plan identically (French CSV)' {
        $dir = Join-Path $script:Temp 'plan-roundtrip'
        [void](Save-EmmPlan -Plan $script:P -Directory $dir -Settings $script:S)
        $back = Import-EmmPlan -Path $dir
        $back.Rows.Count | Should -Be $script:P.Rows.Count
        $a = @($script:P.Rows | Sort-Object ExchangeGuid); $b = @($back.Rows | Sort-Object ExchangeGuid)
        for ($i = 0; $i -lt $a.Count; $i++) { $b[$i].MoveSizeMB | Should -Be $a[$i].MoveSizeMB; $b[$i].Batch | Should -Be $a[$i].Batch; $b[$i].MoveType | Should -Be $a[$i].MoveType }
        $back.Meta.Workload | Should -Be 'All'
        (Find-EmmPlan -OutputPath $script:Temp -Workload User) | Should -BeNullOrEmpty
    }
}

Describe 'Start, complete and cleanup (fictitious organisation)' {
    BeforeAll {
        New-EmmFakeOrg
        foreach ($m in $global:EmmFake.Mailboxes) { $m.FailMove = $false }
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}3' })
        $script:Scope = Resolve-EmmScope -Settings $script:S -Workload User
        $script:I = Get-TestInventory $script:S $script:Scope
        $script:P = New-EmmPlan -Settings $script:S -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
    }
    It 'simulation sends every change with -WhatIf and changes nothing' {
        Reset-EmmActions
        $pre = Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User
        $before = $global:EmmFake.Calls.Count
        Start-EmmMigration -Settings $script:S -Preflight $pre -Simulate 6>$null
        $calls = @($global:EmmFake.Calls | Select-Object -Skip $before)
        $calls.Count | Should -BeGreaterThan 0
        @($calls | Where-Object { -not $_.WhatIf }).Count | Should -Be 0
        $global:EmmFake.Batches.Count | Should -Be 0
    }
    It 'pre-flight: skips a move in progress and a finished move of another origin, replaces a finished move of this tool' {
        $pre = Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User
        @($pre.Items | Where-Object { $_.Row.Alias -eq 'helpdesk.move' })[0].Decision | Should -Be 'Skip'
        $previous = @($pre.Items | Where-Object { $_.Row.Alias -eq 'previous.move' })[0]
        $previous.Decision | Should -Be 'Skip'
        $previous.Reason | Should -BeLike '*not removed by the tool*'
        @($pre.Items | Where-Object { $_.Row.Alias -eq 'retry.me' })[0].RemoveMoveRequest | Should -BeLike '*Retry Me'
        $all = New-TestSettings ([ordered]@{ "(?m)^(\s*ReplaceFinishedMoveRequests\s*=\s*)'[^']*'" = "`${1}'All'" })
        $preAll = Get-EmmStartPreflight -Settings $all -Plan $script:P -Workload User
        @($preAll.Items | Where-Object { $_.Row.Alias -eq 'previous.move' })[0].RemoveMoveRequest | Should -BeLike '*Previous Move'
    }
    It 'creates the batches with MailboxType, and public folders with suspended move requests' {
        Reset-EmmActions
        $pre = Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User
        Start-EmmMigration -Settings $script:S -Preflight $pre 6>$null
        $global:EmmFake.Batches.Count | Should -Be 3
        $rows = @($global:EmmFake.Batches | ForEach-Object { $_.Rows })
        @($rows | Where-Object { $_.EmailAddress -eq 'archive.left@contoso.com' })[0].MailboxType | Should -Be 'ArchiveOnly'
        $pf = @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'New-MoveRequest' -and -not $_.WhatIf })
        $pf.Count | Should -Be 2
        @($pf | Where-Object { -not $_.Parameters['SuspendWhenReadyToComplete'] -or $_.Parameters.ContainsKey('PublicFolder') }).Count | Should -Be 0
        @($global:EmmFake.Calls | Where-Object { $_.Command -eq 'New-MigrationBatch' -and $_.Parameters.ContainsKey('LargeItemLimit') }).Count | Should -Be 0
    }
    It 'starting the same plan again changes nothing' {
        Reset-EmmActions
        $before = $global:EmmFake.Calls.Count
        $pre = Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User
        @($pre.Batches | Where-Object { $_.Action -ne 'Skip' }).Count | Should -Be 0
        Start-EmmMigration -Settings $script:S -Preflight $pre 6>$null
        $global:EmmFake.Calls.Count | Should -Be $before
    }
    It 'completes only synced batches' {
        Reset-EmmActions
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01' 6>$null
        @(Get-EmmActions | Where-Object { $_.Action -eq 'Complete-MigrationBatch' })[0].Status | Should -Be 'Skipped'
        Step-EmmFakeOrg -Rounds 2
        Reset-EmmActions
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01' 6>$null
        @(Get-EmmActions | Where-Object { $_.Action -eq 'Complete-MigrationBatch' })[0].Status | Should -Be 'Success'
    }
    It 'cleanup keeps a synced batch with pending moves and removes a completed one' {
        Step-EmmFakeOrg -Rounds 2
        $plan = Get-EmmCleanupPlan -Settings $script:S
        @($plan.Batches | ForEach-Object { $_.Batch.Name }) | Should -Contain 'Batch01'
        @($plan.KeptBatches | ForEach-Object { $_.Batch.Name }) | Should -Contain 'Batch02'
        @($plan.MoveRequests | Where-Object { $_.BatchName -like 'Helpdesk*' }).Count | Should -Be 0
        Reset-EmmActions
        [void](Invoke-EmmCleanup -Settings $script:S -CleanupPlan $plan 6>$null)
        @($global:EmmFake.Batches | ForEach-Object { $_.Identity }) | Should -Not -Contain 'Batch01'
        $global:EmmFake.ContainsKey('LostMoves') | Should -BeFalse
        @($global:EmmFake.MigrationUsers | Where-Object { $_.Identity -in 'old.user@contoso.com', 'other.tool@contoso.com' }).Count | Should -Be 2
    }
    It 'status joins move requests and migration users' {
        $st = Get-EmmStatus -Settings $script:S
        @($st.Batches | ForEach-Object { $_.Name }) | Should -Contain 'Batch02'
        $users = @($global:EmmFake.MigrationUsers | Where-Object { $_.BatchId -eq 'Batch02' }).Count
        @($st.Rows | Where-Object { $_.Batch -eq 'Batch02' -and $_.ServiceStatus }).Count | Should -Be $users
        @($st.Rows | Where-Object { $_.Batch -eq 'Batch02' -and $_.StatusDetail -eq 'No move request yet' }).Count | Should -Be 0
    }
}

Describe 'Public folder mailboxes alone (no migration batch)' {
    BeforeAll {
        New-EmmFakeOrg
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}2' })
        $script:Scope = Resolve-EmmScope -Settings $script:S -MailboxType PublicFolderMailbox
        $script:I = Get-TestInventory $script:S $script:Scope
        $script:P = New-EmmPlan -Settings $script:S -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
        Reset-EmmActions
        Start-EmmMigration -Settings $script:S -Preflight (Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User) 6>$null
        Step-EmmFakeOrg -Rounds 2
    }
    It 'creates labelled move requests and no batch' {
        $global:EmmFake.Batches.Count | Should -Be 0
        @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -match '^Batch\d\d$' -and $_.Status -eq 'AutoSuspended' }).Count | Should -Be 2
    }
    It 'are followed, completed and cleaned up like batches' {
        $names = @($script:P.Batches | ForEach-Object { $_.Name })
        $names.Count | Should -Be 2
        $managed = @(Get-EmmManagedName -Objects (Get-EmmMigrationObjects -Settings $script:S))
        foreach ($n in $names) { $managed | Should -Contain $n }
        $st = Get-EmmStatus -Settings $script:S
        @($st.Batches | Where-Object { $_.Name -in $names -and $_.Next -like '-Mode Complete*' }).Count | Should -Be 2
        Reset-EmmActions
        Complete-EmmBatch -Settings $script:S -BatchName $names 6>$null
        @(Get-EmmActions | Where-Object { $_.Action -like 'Resume-MoveRequest*' -and $_.Status -eq 'Success' }).Count | Should -Be 2
        Step-EmmFakeOrg -Rounds 2
        $plan = Get-EmmCleanupPlan -Settings $script:S
        @($plan.MoveRequests | Where-Object { $_.ToolBatch -in $names }).Count | Should -Be 2
    }
}

Describe 'Archive policy' {
    It 'refuses to plan archives when the dedicated archive databases do not exist' {
        New-EmmFakeOrg -Small
        $s = New-TestSettings ([ordered]@{ "(?m)^(\s*ArchiveTargetDatabasePattern\s*=\s*)''" = "`${1}'^ARCH-\d+`$`$'" })
        $scope = Resolve-EmmScope -Settings $s -Workload User
        $inv = Get-TestInventory $s $scope
        { New-EmmPlan -Settings $s -DatabaseInventory $inv.Databases -Mailbox $inv.Mailboxes -Scope $scope } | Should -Throw -ExpectedMessage '*ArchiveTargetDatabasePattern*'
    }
    It 'accepts a mapped archive database even when the archive pattern matches nothing' {
        New-EmmFakeOrg -Small
        $s = New-TestSettings ([ordered]@{
                "(?m)^(\s*ArchiveTargetDatabasePattern\s*=\s*)''" = "`${1}'^ARCH-\d+`$`$'"
                "(?m)^(\s*#\s*'DBArchives'\s*=\s*'DB-10')"      = "            'DBArchives' = 'DB-04'"
            })
        $scope = Resolve-EmmScope -Settings $s -MailboxType Archive
        $inv = Get-TestInventory $s $scope
        $p = New-EmmPlan -Settings $s -DatabaseInventory $inv.Databases -Mailbox $inv.Mailboxes -Scope $scope
        @($p.Rows | Where-Object { $_.TargetArchiveDatabase -ne 'DB-04' }).Count | Should -Be 0
        $p.Rows.Count | Should -BeGreaterThan 0
    }
}

Describe 'Ownership and naming' {
    BeforeAll {
        New-EmmFakeOrg -Small
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}1' })
        $script:Scope = Resolve-EmmScope -Settings $script:S -Workload User
        $script:I = Get-TestInventory $script:S $script:Scope
        $script:P = New-EmmPlan -Settings $script:S -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
    }
    It 'never replaces a finished batch that does not belong to the tool' {
        # Same name as the plan batch, but the configuration now uses another prefix for the tool.
        $global:EmmFake.Batches.Add([pscustomobject]@{ Identity = 'Wave01'; Status = 'Completed'; Rows = @(); Created = (Get-Date) })
        $s2 = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchNamePrefix\s*=\s*)'[^']*'" = "`${1}'Wave'"; "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}1' })
        $p2 = New-EmmPlan -Settings $s2 -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
        $s3 = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchNamePrefix\s*=\s*)'[^']*'" = "`${1}'Other'" })
        { Get-EmmStartPreflight -Settings $s3 -Plan $p2 -Workload User } | Should -Throw -ExpectedMessage '*BatchNamePrefix*'
        $p2.Meta['BatchNamePrefix'] = 'Other'
        # Even with matching metadata, rows named after another prefix are refused (the batch-level ownership check stays as a second line).
        { Get-EmmStartPreflight -Settings $s3 -Plan $p2 -Workload User } | Should -Throw -ExpectedMessage '*not a name of this tool*'
        $global:EmmFake.Batches.Clear()
    }
    It 'refuses a plan row whose batch name is not a name of this tool' {
        $p = New-EmmPlan -Settings $script:S -DatabaseInventory $script:I.Databases -Mailbox $script:I.Mailboxes -Scope $script:Scope
        $p.Rows[0].Batch = 'Foreign01'
        { Get-EmmStartPreflight -Settings $script:S -Plan $p -Workload User } | Should -Throw -ExpectedMessage '*not a name of this tool*'
    }
    It 'completes the late public folder moves of a completed batch, and does not suggest the cleanup before' {
        $global:EmmFake.Batches.Clear()
        foreach ($m in $global:EmmFake.Mailboxes) { $m.FailMove = $false }
        Reset-EmmActions
        Start-EmmMigration -Settings $script:S -Preflight (Get-EmmStartPreflight -Settings $script:S -Plan $script:P -Workload User) 6>$null
        $pf = @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq 'Batch01' })[0]
        $pf.Frozen = $true
        Step-EmmFakeOrg -Rounds 2
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01' 6>$null
        Step-EmmFakeOrg -Rounds 2
        @($global:EmmFake.Batches)[0].Status | Should -Be 'Completed'
        $st = Get-EmmStatus -Settings $script:S
        @($st.Batches | Where-Object { $_.Name -eq 'Batch01' })[0].Next | Should -Not -BeLike '*Cleanup*'
        Reset-EmmActions
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01' 6>$null
        @(Get-EmmActions | Where-Object { $_.Detail -like '*not synchronised yet*' }).Count | Should -Be 1
        (Get-EmmCleanupPlan -Settings $script:S).KeptBatches.Count | Should -Be 1
        $pf.Frozen = $false
        Step-EmmFakeOrg -Rounds 2
        @(Get-EmmStatus -Settings $script:S).Batches[0].Next | Should -BeLike '-Mode Complete*'
        Reset-EmmActions
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01' 6>$null
        @(Get-EmmActions | Where-Object { $_.Action -like 'Resume-MoveRequest*' -and $_.Status -eq 'Success' }).Count | Should -Be 1
    }
}

Describe 'On-premises scheduled completion and moving a mailbox again' {
    BeforeAll {
        New-EmmFakeOrg -Small
        foreach ($m in $global:EmmFake.Mailboxes) { $m.FailMove = $false }
        $script:S = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}2' })
        $script:Scope = Resolve-EmmScope -Settings $script:S -Workload User
        $inv = Get-TestInventory $script:S $script:Scope
        $plan = New-EmmPlan -Settings $script:S -DatabaseInventory $inv.Databases -Mailbox $inv.Mailboxes -Scope $script:Scope
        Reset-EmmActions
        Start-EmmMigration -Settings $script:S -Preflight (Get-EmmStartPreflight -Settings $script:S -Plan $plan -Workload User) 6>$null
        Step-EmmFakeOrg -Rounds 2
        Complete-EmmBatch -Settings $script:S -BatchName 'Batch01', 'Batch02' -CompleteAfter (Get-Date).AddHours(2) 6>$null
        Step-EmmFakeOrg
    }
    It 'Cleanup keeps the scheduled batches while their moves wait for the completion time' {
        $c = Get-EmmCleanupPlan -Settings $script:S
        $c.Batches.Count | Should -Be 0
        $c.KeptBatches.Count | Should -Be 2
    }
    It 'Cleanup removes them once their moves are completed, although Exchange leaves them Synced' {
        foreach ($r in $global:EmmFake.MoveRequests) { if ($r.CompleteAfter) { $r.CompleteAfter = (Get-Date).AddMinutes(-1) } }
        Step-EmmFakeOrg -Rounds 2
        @($global:EmmFake.Batches | Where-Object { $_.Status -ne 'Synced' }).Count | Should -Be 0
        @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -like 'MigrationService:*' -and $_.Status -ne 'Completed' }).Count | Should -Be 0
        $c = Get-EmmCleanupPlan -Settings $script:S
        $c.Batches.Count | Should -Be 2
        $c.Batches[0].Reason | Should -BeLike '*CompleteAfter*'
    }
    It 'Start moves a mailbox again without a Cleanup first: earlier batch removed, finished request removed' {
        $user = @($global:EmmFake.MigrationUsers | Where-Object { $_.BatchId -eq 'Batch02' })[0]
        $m = @($global:EmmFake.Mailboxes | Where-Object { $_.ExchangeGuid -eq $user.MailboxGuid })[0]
        $m.Database = 'DB01'; $m.ArchiveDatabase = ''
        $s1 = New-TestSettings ([ordered]@{ "(?m)^(\s*BatchCount\s*=\s*)\d+" = '${1}1' })
        $inv = Get-TestInventory $s1 $script:Scope
        $plan = New-EmmPlan -Settings $s1 -DatabaseInventory $inv.Databases -Mailbox $inv.Mailboxes -Scope $script:Scope
        @($plan.Rows | Where-Object { $_.ExchangeGuid -eq $m.ExchangeGuid }).Count | Should -Be 1
        $pre = Get-EmmStartPreflight -Settings $s1 -Plan $plan -Workload User
        @($pre.Batches)[0].Action | Should -Be 'Replace'
        $item = @($pre.Items | Where-Object { $_.Row.ExchangeGuid -eq $m.ExchangeGuid })[0]
        $item.Decision | Should -Be 'Submit'
        $item.RemoveMoveRequest | Should -Not -BeNullOrEmpty
        $item.PreviousBatch | Should -Be 'Batch02'
        Reset-EmmActions
        Start-EmmMigration -Settings $s1 -Preflight $pre 6>$null
        @(Get-EmmActions | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 0
        @($global:EmmFake.Batches | ForEach-Object { $_.Identity }) -join ',' | Should -Be 'Batch01'
        @($global:EmmFake.Batches[0].Rows | ForEach-Object { $_.EmailAddress }) | Should -Contain $m.PrimarySmtpAddress
        $global:EmmFake.ContainsKey('LostMoves') | Should -BeFalse
    }
}

Describe 'Reports' {
    It 'writes the HTML (marker replaced) and the CSV (BOM, French decimals)' {
        $s = New-TestSettings
        $dir = Join-Path $script:Temp 'report'
        $rows = @([pscustomobject]@{ Name = 'a</script>'; SizeMB = 1.5 }, [pscustomobject]@{ Name = '=cmd'; SizeMB = 2 })
        $files = New-EmmReport -Kind Inventory -Settings $s -Directory $dir -Meta ([ordered]@{ workload = 'All' }) -Tables ([ordered]@{ mailboxes = $rows }) -Csv ([ordered]@{ 'x.csv' = 'mailboxes' })
        $html = [IO.File]::ReadAllText((Join-Path $dir 'Inventory.html'))
        $html | Should -Not -BeLike '*%%DATA%%*'
        $html | Should -Not -BeLike '*a</script>*'
        $bytes = [IO.File]::ReadAllBytes((Join-Path $dir 'x.csv'))
        $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
        $csv = [IO.File]::ReadAllText((Join-Path $dir 'x.csv'))
        $csv | Should -BeLike "*1,5*"
        $csv | Should -BeLike "*'=cmd*"
        @($files | Where-Object { $_.Kind -eq 'CSV' }).Count | Should -Be 1
    }
}

Describe 'Code' {
    It 'every PowerShell file parses' {
        foreach ($f in Get-ChildItem $script:Root -Recurse -Include *.ps1, *.psm1, *.psd1) {
            $errors = $null; [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errors)
            @($errors).Count | Should -Be 0 -Because $f.Name
        }
    }
    It 'code files are ASCII or start with a BOM (Windows PowerShell 5.1 reads them correctly)' {
        foreach ($f in Get-ChildItem $script:Root -Recurse -Include *.ps1, *.psm1, *.psd1) {
            $bytes = [IO.File]::ReadAllBytes($f.FullName)
            $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
            ($bom -or -not @($bytes | Where-Object { $_ -gt 127 }).Count) | Should -BeTrue -Because $f.Name
        }
    }
    It 'the template contains the data marker exactly once' {
        $t = [IO.File]::ReadAllText((Join-Path $script:Root 'templates\Report.template.html'))
        ([regex]::Matches($t, '%%DATA%%')).Count | Should -Be 1
    }
    It 'every exported function exists' {
        $manifest = Import-PowerShellDataFile (Join-Path $script:Root 'ExchangeMailboxMigration.psd1')
        foreach ($name in $manifest.FunctionsToExport) { Get-Command $name -Module ExchangeMailboxMigration -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty -Because $name }
    }
    It 'module and script versions agree' {
        $v = (Import-PowerShellDataFile (Join-Path $script:Root 'ExchangeMailboxMigration.psd1')).ModuleVersion
        foreach ($f in 'ExchangeMailboxMigration.psm1', 'Invoke-ExchangeMailboxMigration.ps1', 'config\ExchangeMailboxMigration.config.psd1', 'templates\Report.template.html') {
            [IO.File]::ReadAllText((Join-Path $script:Root $f)) | Should -BeLike "*Version : $v*" -Because $f
        }
    }
}
