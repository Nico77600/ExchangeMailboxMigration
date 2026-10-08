# Exchange Mailbox Migration

Moves the mailboxes of an Exchange Server organisation from source databases to target databases with migration batches, with system and user mailboxes kept apart and monitoring mailboxes never moved.

This folder contains everything needed to run the tool: Invoke-ExchangeMailboxMigration.ps1, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Exchange Server 2019 / Subscription Edition; Exchange 2016 uses the same cmdlets.
- Windows PowerShell 5.1 only, in the Exchange Management Shell or powershell.exe.
- RBAC roles Mail Recipients, Move Mailboxes, Migration and View-Only Configuration.
- Windows Terminal is recommended for emoji and colours.

## Quick start
```powershell
notepad .\config\ExchangeMailboxMigration.config.psd1      # Databases: source and target patterns

.\Invoke-ExchangeMailboxMigration.ps1                                   # inventory: changes nothing
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan  -Workload System
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System -WhatIf
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan  -Workload User
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload User
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Status -Follow
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Complete -Batch 01,02 -CompleteAfter '2026-10-03 22:00'
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Cleanup
```

## Content
| Item | Role |
|---|---|
| `config\` | Example configuration file. |
| `docs\` | User guide and developer guide, in Markdown and HTML, with images. |
| `templates\` | HTML report template. |
| `ExchangeMailboxMigration.psd1` | PowerShell module manifest. |
| `ExchangeMailboxMigration.psm1` | PowerShell module with the migration functions. |
| `Invoke-ExchangeMailboxMigration.ps1` | Entry script to run. |
| `LICENSE` | MIT license. |
| `README.md` | This package quick start. |

## Documentation
- [User guide](docs/ExchangeMailboxMigration-UserGuide.md) - the path step by step, also `docs/ExchangeMailboxMigration-UserGuide.html`, a single file to open locally
- [Developer guide](docs/ExchangeMailboxMigration-Guide.md) - everything else, also `docs/ExchangeMailboxMigration-Guide.html`, a single file to open locally

Project page, releases and change log: https://github.com/Nico77600/ExchangeMailboxMigration

License: [MIT](LICENSE).
