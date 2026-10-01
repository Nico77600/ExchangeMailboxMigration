#
#  Exchange Mailbox Migration - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 2.0.0
#
#  This file is read by Invoke-ExchangeMailboxMigration.ps1. It is a PowerShell
#  data file: text between quotes, $true / $false, numbers, @( ) for lists and
#  @{ } for tables. Lines starting with # are comments.
#
#  Relative paths (.\reports, .\logs) are relative to the tool folder.
#  Some values can be overridden for one execution on the command line
#  (see "Command-line overrides" in docs\ExchangeMailboxMigration-Guide.md).
#
#  The values below describe the example of the guide: mailboxes moved from the
#  databases DB01..DB09 and DBArchives to the databases DB-01..DB-12, all on
#  Exchange Server 2019. Adapt the 'Databases' section to your organisation and
#  run -Mode Inventory first: it changes nothing and shows how every database
#  and every mailbox is classified.
#
@{
    # ---------------------------------------------------------------------
    # Connection to Exchange (Exchange Management Shell / remote PowerShell)
    # ---------------------------------------------------------------------
    Connection = @{
        # Exchange server used for remote PowerShell (http://<server>/PowerShell).
        # '' = the Exchange cmdlets already loaded (Exchange Management Shell), otherwise
        #      the local server, otherwise the local snap-in.
        ExchangeServer   = ''
        Authentication   = 'Kerberos'   # Kerberos | Negotiate | Default
        # Domain controller passed to New-MoveRequest ('' = Exchange chooses).
        DomainController = ''
    }

    # ---------------------------------------------------------------------
    # Databases: where mailboxes come from and where they go.
    #   A database is a SOURCE when it is a key of DatabaseMap or matches
    #   SourceDatabasePattern. It is a TARGET when it matches TargetDatabasePattern
    #   (or ArchiveTargetDatabasePattern), or is a value of DatabaseMap.
    #   Safety: a target database is never a source, even if it matches both.
    #   Patterns are regular expressions: anchor them with ^ and $.
    # ---------------------------------------------------------------------
    Databases = @{
        SourceDatabasePattern        = '^(DB0[1-9]|DBArchives)$'
        TargetDatabasePattern        = '^DB-(0[1-9]|1[0-2])$'
        # Dedicated databases for archives ('' = archives follow the primary mailbox,
        # or the balancing over TargetDatabasePattern for archive-only moves).
        ArchiveTargetDatabasePattern = ''
        # Fixed mapping source -> target, applied before the balancing.
        # A source database that is not listed is balanced over the target databases.
        DatabaseMap = @{
            # 'DB01'       = 'DB-01'
            # 'DBArchives' = 'DB-10'
        }
        # Balancing of the targets: $true = the current size of each target database
        # is counted (fills the emptiest databases first); $false = planned volume only.
        CountExistingData = $true
    }

    # ---------------------------------------------------------------------
    # Scope: which mailboxes are moved.
    #   Workload System : system mailboxes, moved with individual move requests
    #                     (completed automatically).
    #   Workload User   : user mailboxes and archives, moved with migration batches
    #                     (completion decided by the administrator, -Mode Complete).
    #   Workload All    : both (the plan keeps them apart).
    #   Monitoring mailboxes (HealthMailbox...) are NEVER moved: they are listed in
    #   the inventory as excluded. This cannot be changed.
    # ---------------------------------------------------------------------
    Scope = @{
        Workload           = 'All'      # System | User | All   (command line: -Workload)
        SystemMailboxTypes = @('ArbitrationMailbox', 'AuditLogMailbox', 'AuxAuditLogMailbox', 'DiscoveryMailbox')
        # 'Archive' = the archive mailboxes of the selected types (moved with their primary
        # mailbox, or alone when only the archive is on a source database).
        # 'TeamMailbox' (SharePoint site mailbox, Exchange 2013 or later) can be added; it is
        # listed in the inventory as not selected otherwise.
        UserMailboxTypes   = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox',
                               'LinkedMailbox', 'LinkedRoomMailbox', 'PublicFolderMailbox', 'Archive')
        # Mailboxes never moved (alias, primary SMTP address, name or GUID).
        ExcludeMailboxes   = @()
    }

    # ---------------------------------------------------------------------
    # Plan: how user mailboxes are grouped in migration batches.
    # ---------------------------------------------------------------------
    Plan = @{
        BatchCount      = 12            # Balanced strategy: number of batches (1-99)   (command line: -BatchCount)
        # Balanced          : batches of equal size and equal number of mailboxes
        # PerSourceDatabase : one batch per source database (empties the databases one by one)
        # PerTargetDatabase : one batch per target database
        BatchStrategy   = 'Balanced'
        BatchNamePrefix = 'Batch'       # batches Batch01, Batch02... Use another prefix for another wave.
        MaxPlanAgeDays  = 7             # -Mode Start refuses an older plan (unless -Force)
    }

    # ---------------------------------------------------------------------
    # Move settings (New-MigrationBatch and New-MoveRequest).
    # ---------------------------------------------------------------------
    Move = @{
        BadItemLimit        = 20        # corrupted items tolerated per mailbox
        # Items over the size limit tolerated per mailbox. Applies to the individual move requests (system
        # and public folder mailboxes): Exchange has no such setting for local migration batches.
        LargeItemLimit      = 0
        AcceptLargeDataLoss = $false    # must be $true when a limit is 51 or more (Exchange rule)
        NotificationEmails  = @()       # e-mail addresses that receive the migration batch reports
        # A mailbox with a completed or failed move request cannot be moved again before the request is
        # removed. Which of these FINISHED requests -Mode Start may remove for the planned mailboxes:
        #   Tool : only the requests created by this tool (batches Plan.BatchNamePrefix.., System.BatchName)
        #   All  : also finished requests of another origin (old migrations, other administrators)
        #   None : never; the mailbox is skipped and the request is named in the report
        # A request in progress is never removed: the mailbox is skipped.
        ReplaceFinishedMoveRequests = 'Tool'
    }

    # ---------------------------------------------------------------------
    # System mailboxes (Workload System).
    # ---------------------------------------------------------------------
    System = @{
        BatchName          = 'EMM-SystemMailboxes'  # label of the system move requests (Status, Cleanup)
        TargetDatabase     = ''         # '' = balanced over the target databases, or one database name
        AllowLargeItems    = $true      # large items are copied (replaces LargeItemLimit for these requests)
        WaitForCompletion  = $true      # -Mode Start waits for the system moves to finish
        WaitTimeoutMinutes = 30
        PollSeconds        = 30         # interval between two checks while waiting
    }

    # ---------------------------------------------------------------------
    # Follow-up (-Mode Status -Follow, -Mode Complete -Follow).
    # ---------------------------------------------------------------------
    Status = @{
        RefreshMinutes = 2              # interval of -Follow; the HTML page reloads itself at the same pace
        OpenReport     = $true          # -Follow opens the HTML report in the default browser once
    }

    # ---------------------------------------------------------------------
    # Cleanup (-Mode Cleanup): only objects created by this tool are removed.
    # ---------------------------------------------------------------------
    Cleanup = @{
        IncludeFailed              = $false   # also remove failed / stopped batches and failed move requests
        RemoveOrphanMigrationUsers = $true    # migration users left without their batch
        SettleSeconds              = 15       # wait after removing batches: the migration service forgets their users with a delay
    }

    # ---------------------------------------------------------------------
    # Reports (HTML + CSV), written locally only.
    # ---------------------------------------------------------------------
    Report = @{
        OutputPath   = '.\reports'      # one sub-folder per execution; plans are kept there too
        CsvDelimiter = ';'              # ';' opens directly in Excel with French regional settings
        Organization = ''               # display only, e.g. 'Contoso - Exchange 2019'
    }

    # ---------------------------------------------------------------------
    # Log files (one file per day, deleted after RetentionDays).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 90
    }
}
