---
title: Exchange Mailbox Migration
subtitle: User guide
version: 2.0.0
author: Nicolas Fabert
updated: 2026-10-08
logo: mail
badges: terminal:Windows PowerShell 5.1 only | database:Exchange Server 2016 / 2019 / SE | shield:Monitoring mailboxes never moved
---

# Exchange Mailbox Migration — User guide

> What you need before the first run, then one command per step of the migration: **what is on my databases?**, **where will each mailbox go?**, **start the moves**, **where are they?**, **switch the users over tonight**, **clean up**. The concepts, every configuration key, the reports in detail and the internals are in the [developer guide](ExchangeMailboxMigration-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.
>
> The `Install-Module` commands in this documentation use `-Force`, so they also update or reinstall a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one (as administrator for `-Scope AllUsers`), run `Uninstall-Module <ModuleName> -AllVersions -Force`, then run the `Install-Module` command again.

```cards
checklist | Prerequisites | Chapter 1: Windows PowerShell 5.1, the Exchange roles, and what to check before data moves.
download | Install and configure | Chapter 2: copy the tool to an Exchange server and name the databases emptied and filled.
play | The migration | Chapters 3 to 9: inventory, system mailboxes, plan, start, follow, complete, clean up.
lifebuoy | Results and problems | Chapter 10: the reports, the exit codes, the log. Chapter 11: what to do when something looks wrong.
```

# Part I · Start here

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Exchange | **Exchange Server 2019 / Subscription Edition** (2016 works: same cmdlets). Moves **inside one organisation** only. |
| PowerShell | **Windows PowerShell 5.1 only** (the Exchange Management Shell, or `powershell.exe`). PowerShell 7 is not supported by Microsoft for Exchange Server management: the script stops with an explicit message when started in `pwsh`. |
| Where | An Exchange server (recommended: the Exchange Management Shell is there), or a domain computer that can open remote PowerShell to an Exchange server (`http://<server>/PowerShell`, Kerberos). |
| Modules | None to install: the tool uses the Exchange cmdlets. |
| Permissions | The RBAC roles **Mail Recipients**, **Move Mailboxes**, **Migration** and **View-Only Configuration**. *Organization Management* has them all ([developer guide, chapter 5](ExchangeMailboxMigration-Guide.md#5-prerequisites)). |
| Console | Any. Colours and emoji in Windows Terminal; symbols from the console fonts in the classic console. |

> [!WARNING]
> **Transaction logs.** A move writes every item into the target database, so its transaction logs grow by about the volume moved. Check the free space of the log volumes of the target databases (or enable circular logging for the duration of the migration, if your backup policy allows it) — the plan report gives the volume planned per target database.

Also before data moves: the target databases must hold the volume planned (Plan report, *Databases*: existing + arrivals), and a database being filled has a growing number of logs until the next full backup. Moves are queued by the Mailbox Replication Service of the target servers: large batches simply take longer.

<!-- icon: download -->
## 2. Install and configure

```steps
Build the package | On your workstation, from the repository: `.\tools\New-EmmPackage.ps1`. It writes `..\package\ExchangeMailboxMigration-2.0.0` with only the files needed to run.
Copy it to a server | Zip that folder, copy it to the Exchange server, extract it, then unblock the files.
Describe your databases | In `config\ExchangeMailboxMigration.config.psd1`, section `Databases`: which databases are emptied, which are filled.
Look before moving | `.\Invoke-ExchangeMailboxMigration.ps1` — the inventory changes nothing and shows how every database and every mailbox is classified.
```

```powershell
# On the admin workstation: build the package (only the files needed to run)
.\tools\New-EmmPackage.ps1
# -> ..\package\ExchangeMailboxMigration-2.0.0 : zip it and copy it to the Exchange server

# On the Exchange server
Expand-Archive .\ExchangeMailboxMigration-2.0.0.zip -DestinationPath D:\Tools\ExchangeMailboxMigration
Get-ChildItem D:\Tools\ExchangeMailboxMigration -Recurse -File -Force | Unblock-File   # files downloaded from the internet
cd D:\Tools\ExchangeMailboxMigration
notepad .\config\ExchangeMailboxMigration.config.psd1                   # the Databases section at least
.\Invoke-ExchangeMailboxMigration.ps1                                  # inventory: changes nothing
```

The `Databases` section is the one you must change; everything else has a working default:

```powershell
Databases = @{
    SourceDatabasePattern        = '^(DB0[1-9]|DBArchives)$'   # emptied
    TargetDatabasePattern        = '^DB-(0[1-9]|1[0-2])$'      # filled
    ArchiveTargetDatabasePattern = ''                          # dedicated archive databases ('' = none)
    DatabaseMap = @{ }                                         # fixed mapping, e.g. 'DBArchives' = 'DB-10'
    CountExistingData = $true                                  # balance on existing + planned volume
}
```

Patterns are regular expressions: anchor them (`^…$`), otherwise `DB0` also matches `DB01x`. A target database is never a source, even if it matches both. Every value of the file is checked at start, and all problems are reported together.

| Key you may also set | Default | What it changes |
|---|---|---|
| `Plan.BatchCount` | `12` | Number of batches of the `Balanced` strategy (also `-BatchCount`). |
| `Plan.BatchNamePrefix` | `Batch` | Batch names `Batch01`, `Batch02`… Use another prefix for another wave. |
| `Scope.ExcludeMailboxes` | `@()` | Mailboxes never moved (alias, primary SMTP address, name or GUID). |
| `Move.BadItemLimit` | `20` | Corrupted items tolerated per mailbox. |
| `Report.Organization` | `''` | The name shown in the reports, e.g. `Contoso - Exchange 2019`. |

Every key is described in the [developer guide, chapter 7](ExchangeMailboxMigration-Guide.md#7-configuration). The tool creates `reports\` and `logs\` at the first run.

> [!IMPORTANT]
> Every mode that changes something (**Start**, **Complete**, **Cleanup**) shows what it is going to do and asks for confirmation. `-WhatIf` sends every Exchange command with `-WhatIf`: Exchange checks it, nothing changes. Read-only modes (**Inventory**, **Plan**, **Status**) never change anything.

# Part II · Everyday use

<!-- icon: flow -->
## 3. The path

Run every command from the folder of the tool, in Windows PowerShell 5.1. One mode per step:

```flow
search | Inventory | read only
arrow | |
file | Plan | read only
arrow | review |
play | Start | -WhatIf first
arrow | sync |
refresh | Status | -Follow
arrow | ready |
check | Complete | now or later
arrow | |
wrench | Cleanup | finished objects
```

The whole migration, in the order you run it:

```powershell
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

<!-- icon: search -->
## 4. What is on my databases? (Inventory)

```powershell
.\Invoke-ExchangeMailboxMigration.ps1
```

Read the console and the report: every database with its role (*Source*, *Target*, *Other*) and every mailbox with its decision — *To move*, *On target*, *Not selected*, *Excluded*, *Outside the scope* — and its reason.

![Inventory report: databases and mailboxes with their decision](images/report-inventory.png)

> [!TIP]
> Fix the `Databases` section until the roles are right: a database shown as *Other* is ignored; a database that is both source and target is treated as target. Monitoring mailboxes (`HealthMailbox…`) are listed as *Excluded*: they are never moved, whatever the configuration.

<!-- icon: gear -->
## 5. Move the system mailboxes first

The migration service keeps its data in an arbitration mailbox: move the system mailboxes (arbitration, audit log, auxiliary audit log, discovery) **before** the user batches — or after them, not during. `-Mode Start -Workload User` warns while arbitration mailboxes are still on source databases.

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan  -Workload System
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System -WhatIf
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload System
```

One move request per mailbox, **completed automatically**. The start waits for the moves (30 minutes at most) and reports each result.

![Start of the system workload: individual move requests completed automatically](images/console-start-system.png)

<!-- icon: file -->
## 6. Plan the user mailboxes

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Plan -Workload User
```

Nothing is changed: the plan chooses the target database and the batch of every mailbox and writes `MigrationPlan.csv` and `.json` with its report. Check the batches (volume, types, source → target), the load of each target database (existing, arrivals, departures) and the mailboxes. Click a batch to list its mailboxes.

![Plan report: tiles, batches and target databases](images/report-plan.png)

A plan older than `Plan.MaxPlanAgeDays` (7 days) is refused by Start: make a new one.

<!-- icon: play -->
## 7. Simulate, then start

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload User -WhatIf
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Start -Workload User
```

The **pre-flight** compares the plan with the organisation now and decides per mailbox and per batch; the confirmation question gives the totals. Each batch is a local migration batch whose moves **stop when synchronised**: nobody switches over yet.

![Console: pre-flight, confirmation and submission](images/console-start.png)

| Pre-flight result | Meaning |
|---|---|
| *Create* | the batch is created |
| *Replace* | a batch of the tool with the same name exists and all its moves are finished: it is removed first |
| *Skip* (batch) | a batch with the same name exists and is not finished: follow it, or use another prefix |
| *Move request in progress* | the mailbox is skipped (another move is running) |
| *Moved since the plan* | the mailbox is skipped: create a new plan |

Every decision is listed in the Start report and in `Actions.csv`; the complete table is in the [developer guide, chapter 8](ExchangeMailboxMigration-Guide.md#8-step-by-step).

<!-- icon: refresh -->
## 8. Follow the moves

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Status -Follow
```

`-Follow` rewrites `Status.html` every `Status.RefreshMinutes` (2 by default); the page reloads itself and keeps its filters. It stops when the selected batches are finished. **Ctrl+C stops following, not the migration.**

![Status report: completed, ready to complete, in progress, failed](images/report-status.png)

*Ready to complete* = synchronised (95 %), waiting for completion. A **stalled** flag means the move waits: busy target disk, quarantine… Click a row for the full message.

<!-- icon: check -->
## 9. Complete, then clean up

Completion is the only moment the users notice: their mailbox switches to the target database.

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Complete -Batch 01,02                       # now
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Complete -Batch All -CompleteAfter '2026-10-03 22:00'
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Complete -Batch 03 -Follow                  # and follow it
```

- **Now**: the batch must be *Synced*. The public folder moves of the batch waiting for completion are resumed with it.
- **At a chosen time**: `-CompleteAfter '2026-10-03 22:00'` (local time) schedules every move of the batch. The batch then stays *Synced* in the admin center although its moves complete: `-Mode Status` shows the real state.
- `-Batch` accepts `1`, `03`, `Batch03`, `System` or `All`, and is required with Complete.

![Console: completion of two batches](images/console-complete.png)

```powershell
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Cleanup -WhatIf
.\Invoke-ExchangeMailboxMigration.ps1 -Mode Cleanup
```

Cleanup removes, **for this tool only**: completed batches, completed move requests and migration users left without their batch. A batch with a move not finished is kept, and so is a batch holding failed moves (`-IncludeFailed` removes them once the failures are understood). You do not need Cleanup before moving a mailbox again: Start frees the planned mailboxes itself.

<!-- icon: chart -->
## 10. Results

Every execution writes its folder `reports\<date>_<Mode>[_<Workload>]\` with a self-contained `<Mode>.html` page (it can be sent by e-mail) and its CSV files. The console ends with the folder and the next command to run.

| Report | CSV |
|---|---|
| Inventory | `Inventory-Databases.csv`, `Inventory-Mailboxes.csv` |
| Plan | `MigrationPlan.csv`, `Plan-Batches.csv`, `Plan-Databases.csv` |
| Start / Complete / Cleanup | `Actions.csv` |
| Status | `Status-Batches.csv`, `Status-Mailboxes.csv` |

Every page has a light and a dark theme, a *Next step* box with the command to copy, search and filters, sort by clicking a column, the details of a row and **Export view to CSV**. Sizes include the recoverable items (what a move copies): they are larger than the "mailbox size" of the admin center. CSV files are UTF-8 with BOM, with the `;` delimiter.

![Details of a failed move](images/report-status-details.png)

| Group | What Status shows |
|---|---|
| Completed | `Completed`, `CompletedWithWarning` |
| Ready to complete | `Synced`, `AutoSuspended` |
| In progress | `Queued`, `InProgress`, `CompletionInProgress` |
| Failed / suspended | `Failed`, `Suspended` |

| Exit code | Meaning |
|---|---|
| `0` | Success. |
| `1` | Failure (configuration, connection, unexpected error): see the summary and the log. |
| `2` | Finished with items to look at: skipped or failed mailboxes, batches or removals. |
| `3` | Cancelled at the confirmation question: nothing changed. |

The daily log `logs\ExchangeMailboxMigration_<date>.log` keeps every step, every result and **every change command with its parameters**. Keep it with the reports: together they are the audit trail of the migration.

# Part III · Troubleshoot

<!-- icon: lifebuoy -->
## 11. When something looks wrong

| Message or situation | What to do |
|---|---|
| *Exchange Mailbox Migration runs in Windows PowerShell 5.1 only* | Started in PowerShell 7 (`pwsh`): run it with `powershell.exe` or in the Exchange Management Shell. |
| *Invalid configuration … - …* | Every problem is listed: fix them all in the `.psd1`, run again. |
| *Exchange cmdlets not available to this account: …* | RBAC roles missing (chapter 1); the listed cmdlets tell which. |
| *Cannot open remote PowerShell to …* | Server name, Kerberos (run on a domain computer, use the FQDN), WinRM; or run on the Exchange server in the Exchange Management Shell. |
| *No source database* / *No target database* | The patterns or the map match no database: run the inventory and read the roles. |
| Strange characters in the console | Classic console without UTF-8 font: `$env:EMM_ICONS = 'Ascii'` (or `Symbols`); `NO_COLOR=1` removes colours. |
| A mailbox is *Not selected* | Its type is not in the selection (`Scope.*MailboxTypes`, `-MailboxType`, `-Workload`). |
| A mailbox is *Outside the scope* | It is on a database that is neither source nor target. |
| *The plan is N day(s) old* | Create a new plan (recommended) or use `-Force`. |
| *Still a migration user of the batch X* | That earlier batch still has moves to finish: let it finish or complete it, then start again. |
| *Target database … is not available* | Dismounted or renamed since the plan: mount it or create a new plan. |
| A system move is still running after 30 min | It continues in Exchange: follow it with `-Mode Status -Batch System`. |
| Batch *Syncing* although all its moves are *Synced* | The migration service refreshes the batch status by cycles, sometimes with a long delay: wait and run Status again. |
| *Exchange completes only a Synced batch* | The batch is not synchronised yet: follow it, and complete it once Status shows it *Ready to complete*. |
| A move is *stalled* | Exchange protects the target (disk latency, replication, content indexing): the move resumes by itself. Check the health of the target database copies. |
| A move *Failed* (too many bad items) | Read the message in the Status details. Raise `Move.BadItemLimit` if acceptable, then `-Mode Cleanup -IncludeFailed -Batch NN` and plan and start the mailbox again. |
| The batch stays *Synced* after a scheduled completion | Expected: `-Mode Status` shows the moves completed, `-Mode Cleanup` removes the batch. |

**After the migration.** Run the inventory again: every mailbox should be *On target*, and only monitoring mailboxes remain on the source databases. Run `-Mode Cleanup`. Monitoring mailboxes block `Remove-MailboxDatabase`; they are not data, and Managed Availability recreates them on the remaining databases:

```powershell
Get-Mailbox -Monitoring -Database DB01 | Disable-Mailbox -Confirm:$false
# Recreated automatically on the other databases by the Microsoft Exchange Health Manager service
```

Anything else: [developer guide, Annex A — Troubleshooting](ExchangeMailboxMigration-Guide.md#annex-a--troubleshooting). Recipes for a particular case (only the archives, one database at a time, a forced mapping, retrying a failed mailbox): [developer guide, chapter 9](ExchangeMailboxMigration-Guide.md#9-recipes).
