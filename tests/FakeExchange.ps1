<#
    Exchange Mailbox Migration - fake Exchange organisation for tests and demonstrations.
    Author  : Nicolas Fabert
    Version : 2.0.0

    Defines, in the global scope, the Exchange cmdlets used by the tool (Get-Mailbox, New-MigrationBatch,
    Get-MoveRequestStatistics...) over an in-memory organisation. The objects look like those of remote
    PowerShell: databases and servers as names, sizes as '1.2 GB (1,288,490,188 bytes)'.

    The tool sees these functions as "Exchange cmdlets already loaded" and runs unchanged, so the
    tests and the screenshots of the guide never need an Exchange server.

        . .\tests\FakeExchange.ps1
        New-EmmFakeOrg                 # Contoso: 2 source servers, 2 target servers, ~200 mailboxes
        Step-EmmFakeOrg -Rounds 3      # moves progress (queued -> in progress -> synced / completed)

    Every change command is recorded in $global:EmmFake.Calls (command and parameters).
    The data is fictitious (contoso.com).
#>

function global:New-EmmFakeOrg {
    <# Builds the fictitious organisation. -Small: a few mailboxes only (unit tests). #>
    param([switch]$Small, [int]$Seed = 42)
    $rnd = New-Object System.Random($Seed)
    $state = @{
        Servers = [System.Collections.Generic.List[object]]::new(); Databases = [System.Collections.Generic.List[object]]::new()
        Mailboxes = [System.Collections.Generic.List[object]]::new(); MoveRequests = [System.Collections.Generic.List[object]]::new()
        Batches = [System.Collections.Generic.List[object]]::new(); MigrationUsers = [System.Collections.Generic.List[object]]::new()
        Calls = [System.Collections.Generic.List[object]]::new(); AutoStep = $true; Clock = (Get-Date).AddHours(-2)
    }
    $global:EmmFake = $state
    foreach ($s in 'EX01', 'EX02', 'EX03', 'EX04') { $state.Servers.Add([pscustomobject]@{ Name = $s; AdminDisplayVersion = 'Version 15.2 (Build 1748.10)' }) }
    $dbs = @(
        @('DB01', 'EX01'), @('DB02', 'EX01'), @('DB03', 'EX02'), @('DBArchives', 'EX02'),
        @('DB-01', 'EX03'), @('DB-02', 'EX03'), @('DB-03', 'EX04'), @('DB-04', 'EX04'), @('Mailbox Database 0412345678', 'EX01'))
    foreach ($d in $dbs) { $state.Databases.Add([pscustomobject]@{ Name = $d[0]; Server = $d[1]; Mounted = $true; Recovery = $false; ExtraBytes = [long]0 }) }
    $state.Databases.Add([pscustomobject]@{ Name = 'RDB01'; Server = 'EX02'; Mounted = $false; Recovery = $true; ExtraBytes = [long]0 })

    $new = {
        param([string]$Name, [string]$Type, [string]$Kind, [string]$Db, [string]$ArchiveDb = '', [long]$Bytes = 0, [long]$ArchiveBytes = 0, [string]$Alias = '', [string]$Display = '')
        if (-not $Alias) { $Alias = ($Name -replace '[^A-Za-z0-9.]', '').ToLowerInvariant() }
        if (-not $Display) { $Display = $Name }
        $server = ($state.Databases | Where-Object { $_.Name -eq $Db } | Select-Object -First 1).Server
        $m = [pscustomobject]@{
            Name = $Name; DisplayName = $Display; Alias = $Alias; PrimarySmtpAddress = "$Alias@contoso.com"; RecipientTypeDetails = $Type; Kind = $Kind
            Database = $Db; ArchiveDatabase = $ArchiveDb; Guid = [guid]::NewGuid().ToString(); ExchangeGuid = [guid]::NewGuid().ToString()
            ArchiveGuid = $(if ($ArchiveDb) { [guid]::NewGuid().ToString() } else { '00000000-0000-0000-0000-000000000000' })
            DistinguishedName = "CN=$Name,OU=Mailboxes,DC=contoso,DC=com"; ServerName = $server
            PrimaryBytes = $Bytes; ArchiveBytes = $ArchiveBytes; Items = [long]($Bytes / 90KB); ArchiveItems = [long]($ArchiveBytes / 120KB)
            FailMove = $false; Stall = $false
        }
        $state.Mailboxes.Add($m)
        return $m
    }
    $mb = { param([double]$min, [double]$max) [long]([Math]::Round($min + $rnd.NextDouble() * ($max - $min)) * 1MB) }

    # ---- System mailboxes (on DB01) ----------------------------------------------------------------------------
    foreach ($n in 'SystemMailbox{1f05a927-6a2b-4c8e-9f3d-0b2a3c4d5e6f}', 'SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}', 'SystemMailbox{e0dc1c29-89c3-4034-b678-e6c29d823ed9}',
        'Migration.8f3e7716-2011-43e4-96b1-aba62d229136', 'FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042') {
        [void](& $new -Name $n -Type 'ArbitrationMailbox' -Kind 'Arbitration' -Db 'DB01' -Bytes (& $mb 1 900))
    }
    [void](& $new -Name 'SystemMailbox{8cc370d3-822a-4ab8-a926-bb94bd0641a9}' -Type 'AuditLogMailbox' -Kind 'AuditLog' -Db 'DB01' -Bytes (& $mb 50 400))
    [void](& $new -Name 'SystemMailbox{2ce34405-31be-44f1-a9fa-3a7c2d8f5b10}' -Type 'AuxAuditLogMailbox' -Kind 'AuxAuditLog' -Db 'DB01' -Bytes (& $mb 10 80))
    [void](& $new -Name 'DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}' -Type 'DiscoveryMailbox' -Kind 'Regular' -Db 'DB01' -Bytes (& $mb 1 20) -Display 'Discovery Search Mailbox')

    # ---- Monitoring mailboxes (two per database, never moved) -----------------------------------------------------
    foreach ($d in @($state.Databases | Where-Object { -not $_.Recovery })) {
        for ($i = 0; $i -lt 2; $i++) { [void](& $new -Name ('HealthMailbox' + [guid]::NewGuid().ToString('N')) -Type 'MonitoringMailbox' -Kind 'Monitoring' -Db $d.Name -Bytes (& $mb 1 5)) }
    }

    # ---- User mailboxes --------------------------------------------------------------------------------------------
    $first = 'Alex', 'Camille', 'Dominique', 'Jordan', 'Morgan', 'Sacha', 'Charlie', 'Noa', 'Robin', 'Lou', 'Eden', 'Andrea', 'Claude', 'Maxime', 'Sam'
    $last = 'Martin', 'Bernard', 'Dubois', 'Thomas', 'Robert', 'Richard', 'Petit', 'Durand', 'Leroy', 'Moreau', 'Simon', 'Laurent', 'Lefebvre', 'Michel'
    $count = if ($Small) { 12 } else { 150 }
    $sources = 'DB01', 'DB02', 'DB03'
    for ($i = 0; $i -lt $count; $i++) {
        $f = $first[$i % $first.Count]; $l = $last[[int][Math]::Floor($i / $first.Count) % $last.Count]
        $alias = ('{0}.{1}{2}' -f $f, $l, $(if ($i -ge $first.Count * $last.Count) { $i } else { '' })).ToLowerInvariant()
        $db = $sources[$i % 3]
        $archive = if ($i % 5 -eq 0) { 'DBArchives' } elseif ($i % 23 -eq 0) { $db } else { '' }
        [void](& $new -Name "$f $l" -Type 'UserMailbox' -Kind 'Regular' -Db $db -ArchiveDb $archive -Bytes (& $mb 40 9000) -ArchiveBytes $(if ($archive) { & $mb 200 12000 } else { 0 }) -Alias $alias)
    }
    $shared = 'Accounting', 'Human Resources', 'Service Desk', 'Purchasing', 'Legal', 'Communication', 'Payroll', 'Reception', 'Facilities', 'Security Office', 'Training', 'Archives Office'
    $sharedCount = if ($Small) { 2 } else { $shared.Count }
    for ($i = 0; $i -lt $sharedCount; $i++) { [void](& $new -Name $shared[$i] -Type 'SharedMailbox' -Kind 'Regular' -Db $sources[$i % 3] -Bytes (& $mb 300 20000) -ArchiveDb $(if ($i -eq 0) { 'DBArchives' } else { '' }) -ArchiveBytes $(if ($i -eq 0) { & $mb 5000 30000 } else { 0 })) }
    $rooms = 'Room Paris', 'Room Lyon', 'Room Nantes', 'Room Lille', 'Room Bordeaux', 'Room Marseille'
    $roomCount = if ($Small) { 1 } else { $rooms.Count }
    for ($i = 0; $i -lt $roomCount; $i++) { [void](& $new -Name $rooms[$i] -Type 'RoomMailbox' -Kind 'Regular' -Db $sources[$i % 3] -Bytes (& $mb 5 300)) }
    foreach ($e in @('Projector 1', 'Pool Car 1', 'Pool Car 2') | Select-Object -First $(if ($Small) { 1 } else { 3 })) { [void](& $new -Name $e -Type 'EquipmentMailbox' -Kind 'Regular' -Db 'DB02' -Bytes (& $mb 2 60)) }
    if (-not $Small) {
        [void](& $new -Name 'Partner Liaison' -Type 'LinkedMailbox' -Kind 'Regular' -Db 'DB03' -Bytes (& $mb 100 2000))
        [void](& $new -Name 'Project Team Site' -Type 'TeamMailbox' -Kind 'Regular' -Db 'DB02' -Bytes (& $mb 50 500))
    }
    [void](& $new -Name 'PF-Root' -Type 'PublicFolderMailbox' -Kind 'PublicFolder' -Db 'DB02' -Bytes (& $mb 500 4000))
    if (-not $Small) { [void](& $new -Name 'PF-Archive2019' -Type 'PublicFolderMailbox' -Kind 'PublicFolder' -Db 'DB03' -Bytes (& $mb 2000 8000)) }

    # ---- Special cases ------------------------------------------------------------------------------------------------
    [void](& $new -Name 'Already Moved' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB-02' -Bytes (& $mb 100 900) -Alias 'already.moved')
    [void](& $new -Name 'Archive Left Behind' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB-01' -ArchiveDb 'DBArchives' -Bytes (& $mb 300 900) -ArchiveBytes (& $mb 2000 6000) -Alias 'archive.left')
    [void](& $new -Name 'Archive Already Moved' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB02' -ArchiveDb 'DB-03' -Bytes (& $mb 300 900) -ArchiveBytes (& $mb 1000 3000) -Alias 'archive.moved')
    [void](& $new -Name 'Other Database User' -Type 'UserMailbox' -Kind 'Regular' -Db 'Mailbox Database 0412345678' -Bytes (& $mb 100 900) -Alias 'other.db')
    [void](& $new -Name 'Chief Executive' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB01' -Bytes (& $mb 8000 9000) -Alias 'ceo')
    $done = & $new -Name 'Previous Move' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB03' -Bytes (& $mb 100 900) -Alias 'previous.move'
    $busy = & $new -Name 'Helpdesk Move' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB02' -Bytes (& $mb 100 900) -Alias 'helpdesk.move'
    $state.MoveRequests.Add((New-EmmFakeMove -Mailbox $done -BatchName 'Helpdesk-2025' -Target 'DB03' -Status 'Completed'))
    $busyMove = New-EmmFakeMove -Mailbox $busy -BatchName 'Helpdesk-2026' -Target 'DB02' -Status 'InProgress'
    $busyMove.Frozen = $true
    $state.MoveRequests.Add($busyMove)
    # A failed move of an earlier wave of this tool (its batch already removed): replaced at the next start.
    $retry = & $new -Name 'Retry Me' -Type 'UserMailbox' -Kind 'Regular' -Db 'DB01' -Bytes (& $mb 100 900) -Alias 'retry.me'
    $state.MoveRequests.Add((New-EmmFakeMove -Mailbox $retry -BatchName 'MigrationService:Batch09' -Target 'DB-01' -Status 'Failed'))
    # Migration users of other origins, without their batch: never removed by the tool.
    $state.MigrationUsers.Add([pscustomobject]@{ Identity = 'old.user@contoso.com'; BatchId = ''; Status = 'Failed'; MailboxGuid = [guid]::NewGuid().ToString(); LastSync = $null; Error = '' })
    $state.MigrationUsers.Add([pscustomobject]@{ Identity = 'other.tool@contoso.com'; BatchId = 'OtherTool-01'; Status = 'Completed'; MailboxGuid = [guid]::NewGuid().ToString(); LastSync = $null; Error = '' })
    if (-not $Small) {
        ($state.Mailboxes | Where-Object { $_.Alias -eq 'robin.bernard' }).FailMove = $true
        ($state.Mailboxes | Where-Object { $_.Alias -eq 'eden.bernard' }).Stall = $true
    }
}

function global:New-EmmFakeMove {
    param($Mailbox, [string]$BatchName, [string]$Target, [string]$ArchiveTarget = '', [string]$Status = 'Queued', [bool]$Suspend = $false, [string]$MailboxType = '')
    return [pscustomobject]@{
        Identity = "contoso.com/Mailboxes/$($Mailbox.Name)"; DisplayName = $Mailbox.DisplayName; Alias = $Mailbox.Alias; ExchangeGuid = $Mailbox.ExchangeGuid
        Mailbox = $Mailbox; Status = $Status; StatusDetail = ''; BatchName = $BatchName; SourceDatabase = $Mailbox.Database; TargetDatabase = $Target
        TargetArchiveDatabase = $ArchiveTarget; MailboxType = $MailboxType; Suspend = $Suspend; Percent = $(if ($Status -like 'Completed*') { 100 } else { 0 })
        CompleteAfter = $null; Message = ''; FailureType = ''; LastUpdate = (Get-Date); Frozen = $false
    }
}

function global:Write-EmmFakeCall {
    param([string]$Command, [hashtable]$Parameters, [bool]$WhatIf)
    $text = foreach ($k in @($Parameters.Keys | Where-Object { $_ -notin 'ErrorAction', 'WarningAction', 'Confirm', 'WhatIf', 'WarningVariable' } | Sort-Object)) {
        $v = $Parameters[$k]; if ($v -is [byte[]]) { "-$k <csv>" } elseif ($v -is [switch] -or $v -is [bool]) { "-$k" } else { "-$k $v" }
    }
    $global:EmmFake.Calls.Add([pscustomobject]@{ Command = $Command; Parameters = $Parameters; WhatIf = $WhatIf; Text = "$Command $($text -join ' ')" })
}

function global:Format-EmmFakeSize([long]$Bytes) {
    $gb = $Bytes / 1GB
    $label = if ($gb -ge 1) { '{0:0.##} GB' -f $gb } else { '{0:0.##} MB' -f ($Bytes / 1MB) }
    return '{0} ({1} bytes)' -f $label, $Bytes.ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function global:Find-EmmFakeMailbox([string]$Identity) {
    return @($global:EmmFake.Mailboxes | Where-Object { $Identity -in $_.DistinguishedName, $_.Guid, $_.ExchangeGuid, $_.Alias, $_.PrimarySmtpAddress, $_.Name })[0]
}

function global:ConvertTo-EmmFakeMailbox($m) {
    return [pscustomobject]@{
        Name = $m.Name; DisplayName = $m.DisplayName; Alias = $m.Alias; PrimarySmtpAddress = $m.PrimarySmtpAddress; RecipientTypeDetails = $m.RecipientTypeDetails
        Database = $m.Database; ArchiveDatabase = $(if ($m.ArchiveDatabase) { $m.ArchiveDatabase } else { $null }); Guid = $m.Guid; ExchangeGuid = $m.ExchangeGuid
        ArchiveGuid = $m.ArchiveGuid; DistinguishedName = $m.DistinguishedName; ServerName = $m.ServerName; Identity = "contoso.com/Mailboxes/$($m.Name)"
    }
}

# ---- Read cmdlets -------------------------------------------------------------------------------------------------------------

function global:Get-ExchangeServer { [CmdletBinding()] param() $global:EmmFake.Servers | ForEach-Object { [pscustomobject]@{ Name = $_.Name; AdminDisplayVersion = $_.AdminDisplayVersion; ServerRole = 'Mailbox' } } }

function global:Get-MailboxDatabase {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [switch]$Status)
    foreach ($d in @($global:EmmFake.Databases | Where-Object { -not $Identity -or $_.Name -eq $Identity })) {
        $bytes = [long](@($global:EmmFake.Mailboxes | Where-Object { $_.Database -eq $d.Name } | Measure-Object PrimaryBytes -Sum)[0].Sum) + [long](@($global:EmmFake.Mailboxes | Where-Object { $_.ArchiveDatabase -eq $d.Name } | Measure-Object ArchiveBytes -Sum)[0].Sum)
        $o = [pscustomobject]@{ Name = $d.Name; Server = $d.Server; Recovery = $d.Recovery }
        if ($Status) {
            $o | Add-Member Mounted $d.Mounted
            $o | Add-Member DatabaseSize (Format-EmmFakeSize ([long]($bytes * 1.15) + 256MB))
            $o | Add-Member AvailableNewMailboxSpace (Format-EmmFakeSize 512MB)
        }
        $o
    }
}

function global:Get-Mailbox {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Identity, [switch]$Arbitration, [switch]$AuditLog, [switch]$AuxAuditLog, [switch]$Monitoring, [switch]$PublicFolder,
        [string]$RecipientTypeDetails, [object]$ResultSize, [string]$Database)
    $kind = if ($Arbitration) { 'Arbitration' } elseif ($AuditLog) { 'AuditLog' } elseif ($AuxAuditLog) { 'AuxAuditLog' } elseif ($Monitoring) { 'Monitoring' } elseif ($PublicFolder) { 'PublicFolder' } else { 'Regular' }
    $list = @($global:EmmFake.Mailboxes | Where-Object { $_.Kind -eq $kind })
    if ($RecipientTypeDetails) { $list = @($list | Where-Object { $_.RecipientTypeDetails -eq $RecipientTypeDetails }) }
    if ($Identity) { $m = Find-EmmFakeMailbox $Identity; $list = @($list | Where-Object { $_ -eq $m }) }
    if ($Database) { $list = @($list | Where-Object { $_.Database -eq $Database }) }
    foreach ($m in $list) { ConvertTo-EmmFakeMailbox $m }
}

function global:Get-MailboxStatistics {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Database, [switch]$Archive)
    $out = @()
    foreach ($m in $global:EmmFake.Mailboxes) {
        if ($Database) {
            if ($m.Database -eq $Database -and $m.PrimaryBytes) { $out += , @($m, $false) }
            if ($m.ArchiveDatabase -eq $Database -and $m.ArchiveBytes) { $out += , @($m, $true) }
        } elseif ($Identity -and $m -eq (Find-EmmFakeMailbox $Identity)) { $out += , @($m, [bool]$Archive) }
    }
    foreach ($pair in $out) {
        $m = $pair[0]; $isArchive = $pair[1]; $bytes = if ($isArchive) { $m.ArchiveBytes } else { $m.PrimaryBytes }
        [pscustomobject]@{
            DisplayName = $m.DisplayName; MailboxGuid = $(if ($isArchive) { $m.ArchiveGuid } else { $m.ExchangeGuid }); IsArchiveMailbox = $isArchive
            TotalItemSize = Format-EmmFakeSize ([long]($bytes * 0.92)); TotalDeletedItemSize = Format-EmmFakeSize ([long]($bytes * 0.08))
            ItemCount = $(if ($isArchive) { $m.ArchiveItems } else { $m.Items }); DisconnectReason = $null; Database = $(if ($isArchive) { $m.ArchiveDatabase } else { $m.Database })
        }
    }
}

function global:Get-Recipient { [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity) $m = Find-EmmFakeMailbox $Identity; if ($m) { ConvertTo-EmmFakeMailbox $m } }

function global:Get-MoveRequest {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$BatchName, [object]$ResultSize)
    if ($global:EmmFake.AutoStep) { Step-EmmFakeOrg -IndividualOnly }
    foreach ($r in @($global:EmmFake.MoveRequests | Where-Object { (-not $BatchName -or $_.BatchName -eq $BatchName) -and (-not $Identity -or $_.Identity -eq $Identity -or $_.Mailbox -eq (Find-EmmFakeMailbox $Identity)) })) {
        [pscustomobject]@{ Identity = $r.Identity; DisplayName = $r.DisplayName; Alias = $r.Alias; ExchangeGuid = $r.ExchangeGuid; Status = $r.Status; BatchName = $r.BatchName; TargetDatabase = $r.TargetDatabase; SourceDatabase = $r.SourceDatabase }
    }
}

function global:Get-MoveRequestStatistics {
    [CmdletBinding()] param([Parameter(Position = 0, ValueFromPipelineByPropertyName)][string]$Identity)
    process {
        foreach ($r in @($global:EmmFake.MoveRequests | Where-Object { $_.Identity -eq $Identity })) {
            $m = $r.Mailbox
            $primary = $r.MailboxType -ne 'ArchiveOnly'; $archive = $r.MailboxType -in '', 'PrimaryAndArchive', 'ArchiveOnly' -and $m.ArchiveDatabase
            $total = [long]$(if ($primary) { $m.PrimaryBytes } else { 0 }); $arch = [long]$(if ($archive) { $m.ArchiveBytes } else { 0 })
            [pscustomobject]@{
                DisplayName = $r.DisplayName; Alias = $r.Alias; ExchangeGuid = $r.ExchangeGuid; Status = $r.Status; StatusDetail = $r.StatusDetail; PercentComplete = $r.Percent
                TotalMailboxSize = Format-EmmFakeSize $total; TotalArchiveSize = Format-EmmFakeSize $arch; BytesTransferred = Format-EmmFakeSize ([long](($total + $arch) * $r.Percent / 100))
                SourceDatabase = $r.SourceDatabase; TargetDatabase = $r.TargetDatabase; TargetArchiveDatabase = $r.TargetArchiveDatabase; BatchName = $r.BatchName
                Message = $r.Message; FailureType = $r.FailureType; LastUpdateTimestamp = $r.LastUpdate; CompleteAfter = $r.CompleteAfter
            }
        }
    }
}

function global:Get-MigrationBatch {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity)
    foreach ($b in @($global:EmmFake.Batches | Where-Object { -not $Identity -or $_.Identity -eq $Identity })) {
        $users = @($global:EmmFake.MigrationUsers | Where-Object { $_.BatchId -eq $b.Identity })
        [pscustomobject]@{
            Identity = $b.Identity; Status = $b.Status; TotalCount = $users.Count; SyncedCount = @($users | Where-Object { $_.Status -eq 'Synced' }).Count
            FinalizedCount = @($users | Where-Object { $_.Status -eq 'Completed' }).Count; FailedCount = @($users | Where-Object { $_.Status -eq 'Failed' }).Count; CreationDateTime = $b.Created
        }
    }
}

function global:Get-MigrationUser {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$BatchId, [object]$ResultSize)
    foreach ($u in @($global:EmmFake.MigrationUsers | Where-Object { (-not $BatchId -or $_.BatchId -eq $BatchId) -and (-not $Identity -or $_.Identity -eq $Identity) })) {
        [pscustomobject]@{ Identity = $u.Identity; MailboxEmailAddress = $u.Identity; BatchId = $u.BatchId; Status = $u.Status; MailboxGuid = $u.MailboxGuid; LastSuccessfulSyncTime = $u.LastSync; ErrorSummary = $u.Error; RecipientType = 'UserMailbox' }
    }
}

function global:Get-MigrationUserStatistics {
    [CmdletBinding()] param([Parameter(Position = 0, ValueFromPipelineByPropertyName)][string]$Identity)
    process { Get-MigrationUser -Identity $Identity | ForEach-Object { $_ | Add-Member Error $_.ErrorSummary -PassThru } }
}

# ---- Change cmdlets ---------------------------------------------------------------------------------------------------------------

function global:New-MoveRequest {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Position = 0, Mandatory)][string]$Identity, [string]$TargetDatabase, [string]$ArchiveTargetDatabase, [switch]$PrimaryOnly, [switch]$ArchiveOnly,
        [string]$BatchName, [object]$BadItemLimit, [object]$LargeItemLimit, [switch]$AllowLargeItems, [switch]$AcceptLargeDataLoss, [string]$DomainController, [switch]$SuspendWhenReadyToComplete)
    Write-EmmFakeCall 'New-MoveRequest' $PSBoundParameters ([bool]$WhatIfPreference)
    $m = Find-EmmFakeMailbox $Identity
    if (-not $m) { throw "The operation couldn't be performed because object '$Identity' couldn't be found." }
    if ($m.Kind -eq 'Monitoring') { throw "FAKE GUARD: monitoring mailbox $($m.Name) must never be moved." }
    if ($AllowLargeItems -and $PSBoundParameters.ContainsKey('LargeItemLimit')) { throw 'Parameter set cannot be resolved using the specified named parameters (AllowLargeItems with LargeItemLimit).' }
    if ([int]"$BadItemLimit" -ge 51 -and -not $AcceptLargeDataLoss) { throw 'BadItemLimit 51 or more requires AcceptLargeDataLoss.' }
    foreach ($t in @($TargetDatabase, $ArchiveTargetDatabase) | Where-Object { $_ }) { if (-not ($global:EmmFake.Databases | Where-Object { $_.Name -eq $t })) { throw "Couldn't find database '$t'." } }
    if ($global:EmmFake.MoveRequests | Where-Object { $_.Mailbox -eq $m }) { throw "User '$($m.Name)' already has a move request." }
    if ($PSCmdlet.ShouldProcess($m.Name, 'Create move request')) {
        $type = if ($PrimaryOnly) { 'PrimaryOnly' } elseif ($ArchiveOnly) { 'ArchiveOnly' } else { '' }
        $global:EmmFake.MoveRequests.Add((New-EmmFakeMove -Mailbox $m -BatchName $BatchName -Target $TargetDatabase -ArchiveTarget $ArchiveTargetDatabase -Suspend ([bool]$SuspendWhenReadyToComplete) -MailboxType $type))
    }
}

function global:Set-MoveRequest {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity, [datetime]$CompleteAfter)
    Write-EmmFakeCall 'Set-MoveRequest' $PSBoundParameters ([bool]$WhatIfPreference)
    $r = @($global:EmmFake.MoveRequests | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $r) { throw "Move request '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Set move request')) { $r.CompleteAfter = $CompleteAfter }
}

function global:Resume-MoveRequest {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity)
    Write-EmmFakeCall 'Resume-MoveRequest' $PSBoundParameters ([bool]$WhatIfPreference)
    $r = @($global:EmmFake.MoveRequests | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $r) { throw "Move request '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Resume move request')) { $r.Suspend = $false; if ($r.Status -in 'AutoSuspended', 'Synced' -and -not ($r.CompleteAfter -and $r.CompleteAfter -gt (Get-Date))) { $r.Status = 'CompletionInProgress' } }
}

function global:Remove-MoveRequest {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity)
    Write-EmmFakeCall 'Remove-MoveRequest' $PSBoundParameters ([bool]$WhatIfPreference)
    $r = @($global:EmmFake.MoveRequests | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $r) { throw "Move request '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Remove move request')) { [void]$global:EmmFake.MoveRequests.Remove($r) }
}

function global:New-MigrationBatch {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Name, [switch]$Local, [byte[]]$CSVData, [object]$BadItemLimit, [string[]]$NotificationEmails)
    Write-EmmFakeCall 'New-MigrationBatch' $PSBoundParameters ([bool]$WhatIfPreference)
    if (-not $Local) { throw 'Only local batches are emulated.' }
    if ($global:EmmFake.Batches | Where-Object { $_.Identity -eq $Name }) { throw "A migration batch named '$Name' already exists." }
    $rows = @([Text.Encoding]::UTF8.GetString($CSVData) | ConvertFrom-Csv)
    foreach ($r in $rows) {
        $m = Find-EmmFakeMailbox $r.EmailAddress
        if (-not $m) { throw "The user '$($r.EmailAddress)' couldn't be found." }
        if ($m.Kind -eq 'Monitoring') { throw "FAKE GUARD: monitoring mailbox $($m.Name) must never be moved." }
        if ($r.MailboxType -and $r.MailboxType -notin 'PrimaryOnly', 'ArchiveOnly', 'PrimaryAndArchive') { throw "Invalid MailboxType '$($r.MailboxType)'." }
        if ($r.MailboxType -eq 'ArchiveOnly' -and $r.TargetDatabase) { throw 'ArchiveOnly rows must not have a TargetDatabase in this emulation.' }
        if ($global:EmmFake.MigrationUsers | Where-Object { $_.MailboxGuid -eq $m.ExchangeGuid }) { throw "The user '$($r.EmailAddress)' already exists in another migration batch." }
    }
    if ($PSCmdlet.ShouldProcess($Name, 'Create migration batch')) {
        $global:EmmFake.Batches.Add([pscustomobject]@{ Identity = $Name; Status = 'Created'; Rows = $rows; Created = (Get-Date) })
        foreach ($r in $rows) { $m = Find-EmmFakeMailbox $r.EmailAddress; $global:EmmFake.MigrationUsers.Add([pscustomobject]@{ Identity = $m.PrimarySmtpAddress; BatchId = $Name; Status = 'Queued'; MailboxGuid = $m.ExchangeGuid; LastSync = $null; Error = '' }) }
    }
}

function global:Start-MigrationBatch {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity)
    Write-EmmFakeCall 'Start-MigrationBatch' $PSBoundParameters ([bool]$WhatIfPreference)
    $b = @($global:EmmFake.Batches | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $b) { throw "Migration batch '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Start migration batch')) {
        $b.Status = 'Syncing'
        foreach ($r in $b.Rows) {
            $m = Find-EmmFakeMailbox $r.EmailAddress
            $global:EmmFake.MoveRequests.Add((New-EmmFakeMove -Mailbox $m -BatchName "MigrationService:$Identity" -Target $r.TargetDatabase -ArchiveTarget $r.TargetArchiveDatabase -Suspend $true -MailboxType $r.MailboxType))
        }
    }
}

function global:Complete-MigrationBatch {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity)
    Write-EmmFakeCall 'Complete-MigrationBatch' $PSBoundParameters ([bool]$WhatIfPreference)
    $b = @($global:EmmFake.Batches | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $b) { throw "Migration batch '$Identity' not found." }
    if ($b.Status -ne 'Synced') { throw "The migration batch '$Identity' can't be completed because its status is '$($b.Status)'." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Complete migration batch')) {
        $b.Status = 'Completing'
        foreach ($r in @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq "MigrationService:$Identity" -and $_.Status -in 'Synced', 'AutoSuspended' })) { $r.Suspend = $false; $r.Status = 'CompletionInProgress' }
    }
}

function global:Remove-MigrationBatch {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity, [switch]$Force)
    Write-EmmFakeCall 'Remove-MigrationBatch' $PSBoundParameters ([bool]$WhatIfPreference)
    $b = @($global:EmmFake.Batches | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $b) { throw "Migration batch '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Remove migration batch')) {
        [void]$global:EmmFake.Batches.Remove($b)
        foreach ($u in @($global:EmmFake.MigrationUsers | Where-Object { $_.BatchId -eq $Identity })) { [void]$global:EmmFake.MigrationUsers.Remove($u) }
        # Removing a batch cancels its moves that are not completed (the synchronised data is lost).
        foreach ($r in @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq "MigrationService:$Identity" -and $_.Status -notlike 'Completed*' })) { [void]$global:EmmFake.MoveRequests.Remove($r); $global:EmmFake.LostMoves++ }
    }
}

function global:Remove-MigrationUser {
    [CmdletBinding(SupportsShouldProcess)] param([Parameter(Position = 0, Mandatory)][string]$Identity)
    Write-EmmFakeCall 'Remove-MigrationUser' $PSBoundParameters ([bool]$WhatIfPreference)
    $u = @($global:EmmFake.MigrationUsers | Where-Object { $_.Identity -eq $Identity })[0]
    if (-not $u) { throw "Migration user '$Identity' not found." }
    if ($PSCmdlet.ShouldProcess($Identity, 'Remove migration user')) { [void]$global:EmmFake.MigrationUsers.Remove($u) }
}

# ---- Time -------------------------------------------------------------------------------------------------------------------------

function global:Step-EmmFakeOrg {
    <#
    .SYNOPSIS
        Moves progress by one stage per round: Queued -> InProgress -> Synced (batch / suspended requests)
        or Completed; CompletionInProgress -> Completed (the mailbox is then on its target database).
        -IndividualOnly: only the move requests outside migration batches (system moves).
    #>
    param([int]$Rounds = 1, [switch]$IndividualOnly)
    for ($i = 0; $i -lt $Rounds; $i++) {
        foreach ($r in $global:EmmFake.MoveRequests.ToArray()) {
            if ($r.Frozen -or ($IndividualOnly -and ($r.BatchName -like 'MigrationService:*' -or $r.Suspend))) { continue }
            $r.LastUpdate = (Get-Date)
            switch ($r.Status) {
                'Queued' { $r.Status = 'InProgress'; $r.Percent = 35 }
                'InProgress' {
                    if ($r.Mailbox.FailMove) { $r.Status = 'Failed'; $r.Percent = 40; $r.FailureType = 'TooManyBadItemsPermanentException'; $r.Message = 'Error: This mailbox exceeded the maximum number of corrupted items that were specified for this move request.' }
                    elseif ($r.Suspend) { $r.Status = $(if ($r.BatchName -like 'MigrationService:*') { 'Synced' } else { 'AutoSuspended' }); $r.Percent = 95; if ($r.Mailbox.Stall) { $r.StatusDetail = 'StalledDueToTarget_DiskLatency' } }
                    else { $r.Status = 'CompletionInProgress'; $r.Percent = 98 }
                }
                { $_ -in 'Synced', 'AutoSuspended' } { if (-not $r.Suspend -and -not ($r.CompleteAfter -and $r.CompleteAfter -gt (Get-Date))) { $r.Status = 'CompletionInProgress'; $r.Percent = 98 } }
                'CompletionInProgress' {
                    $r.Status = 'Completed'; $r.Percent = 100; $r.StatusDetail = ''
                    if ($r.MailboxType -ne 'ArchiveOnly' -and $r.TargetDatabase) { $r.Mailbox.Database = $r.TargetDatabase }
                    if ($r.MailboxType -ne 'PrimaryOnly' -and $r.Mailbox.ArchiveDatabase -and $r.TargetArchiveDatabase) { $r.Mailbox.ArchiveDatabase = $r.TargetArchiveDatabase }
                }
            }
        }
        foreach ($b in $global:EmmFake.Batches.ToArray()) {
            $moves = @($global:EmmFake.MoveRequests | Where-Object { $_.BatchName -eq "MigrationService:$($b.Identity)" })
            foreach ($u in @($global:EmmFake.MigrationUsers | Where-Object { $_.BatchId -eq $b.Identity })) {
                $mr = @($moves | Where-Object { $_.ExchangeGuid -eq $u.MailboxGuid })[0]
                if ($mr) { $u.Status = switch ($mr.Status) { 'Completed' { 'Completed' } 'Synced' { 'Synced' } 'Failed' { 'Failed' } 'CompletionInProgress' { 'Completing' } default { 'Syncing' } }; $u.Error = $mr.Message; if ($mr.Status -in 'Synced', 'Completed') { $u.LastSync = Get-Date } }
            }
            if (-not $moves.Count) { continue }
            $open = @($moves | Where-Object { $_.Status -notin 'Synced', 'Completed', 'Failed' })
            if (-not @($moves | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count -and $b.Status -in 'Completing', 'Synced' -and @($moves | Where-Object { $_.Status -eq 'Completed' }).Count) { $b.Status = $(if (@($moves | Where-Object { $_.Status -eq 'Failed' }).Count) { 'CompletedWithErrors' } else { 'Completed' }) }
            elseif (-not $open.Count -and $b.Status -eq 'Syncing') { $b.Status = 'Synced' }
        }
    }
}
