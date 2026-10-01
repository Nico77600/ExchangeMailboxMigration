#
#  Exchange Mailbox Migration - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-ExchangeMailboxMigration.ps1 (Import-Module by path).
#
@{
    RootModule        = 'ExchangeMailboxMigration.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = '5b0e3c1a-7d42-4f6e-9a1b-2c8d4e6f7a90'
    Author            = 'Nicolas Fabert'
    Description       = 'Exchange Mailbox Migration: moves Exchange Server mailboxes from source to target databases with migration batches - system and user mailboxes kept apart, monitoring mailboxes never moved - with HTML/CSV reports.'
    PowerShellVersion = '5.1'
    # Windows PowerShell only: PowerShell 7 is not supported by Microsoft for Exchange Server management.
    CompatiblePSEditions = @('Desktop')

    # Functions called by Invoke-ExchangeMailboxMigration.ps1 and by the tests. The other functions stay
    # internal to the module: add a function here only when the script or a test calls it.
    FunctionsToExport = @(
        'Import-EmmConfiguration', 'Resolve-EmmScope'
        'Start-EmmLog', 'Stop-EmmLog', 'Write-EmmLog'
        'Write-EmmBanner', 'Write-EmmStep', 'Write-EmmItem', 'Write-EmmTable', 'Write-EmmSummary'
        'Format-EmmNumber', 'Format-EmmDuration', 'Format-EmmSize', 'ConvertTo-EmmMB', 'ConvertTo-EmmJson'
        'Connect-EmmExchange', 'Disconnect-EmmExchange'
        'Get-EmmDatabaseClassification', 'Get-EmmDatabaseInventory', 'Get-EmmMailboxInventory', 'Get-EmmMigrationObjects', 'Get-EmmManagedName'
        'New-EmmPlan', 'Save-EmmPlan', 'Import-EmmPlan', 'Find-EmmPlan'
        'Reset-EmmActions', 'Get-EmmActions', 'Get-EmmStartPreflight', 'Start-EmmMigration', 'Complete-EmmBatch', 'Get-EmmCleanupPlan', 'Invoke-EmmCleanup'
        'Get-EmmStatus', 'Select-EmmBatchName', 'Test-EmmMoveType', 'Get-EmmBar', 'Format-EmmDate'
        'New-EmmRunDirectory', 'New-EmmReport', 'Export-EmmCsv'
        'Get-EmmIconSet', 'Set-EmmMailboxSelection', 'ConvertTo-EmmMailboxRecord'
    )
}
