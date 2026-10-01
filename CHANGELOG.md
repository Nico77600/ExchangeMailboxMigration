# Changelog — Exchange Mailbox Migration

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the guide, Annex D).
Author: Nicolas Fabert.

## [2.0.0] — 2026-10-01

First public release. Rewrite of the internal script `Migrate-ExchangeBals` 1.3/1.4 (steps 18 to 22 of an Exchange 2019 deployment framework), aligned with the standards of the author's other tools.

### Added
- **One entry point**, `Invoke-ExchangeMailboxMigration.ps1`, with one mode per step: `Inventory`, `Plan`, `Start`, `Status`, `Complete`, `Cleanup`. Comment-based help, parameter checks per mode, exit codes 0 / 1 / 2 / 3.
- **Configuration file** `config\ExchangeMailboxMigration.config.psd1` (connection, databases, scope, plan, moves, system mailboxes, follow-up, cleanup, reports, logs), checked at start with every problem reported at once; a few command-line overrides for one execution.
- **Workloads kept apart**: `-Workload System` (arbitration, audit log, auxiliary audit log, discovery: individual move requests labelled `EMM-SystemMailboxes`, completed automatically, waited for) and `-Workload User` (user, shared, room, equipment, linked, public folder mailboxes and archives: local migration batches completed on demand). `-MailboxType` restricts a plan to some types, `Archive` included.
- **Monitoring mailboxes never moved**, checked when reading, in the plan and before each submission; a configuration that selects them is refused. Listed as excluded in the inventory.
- **Inventory**: every database with its role (source / target / archive target / other), every mailbox with its decision and reason (*To move*, *On target*, *Not selected*, *Excluded*, *Outside the scope*).
- **Plan** reviewed before anything happens: one row per mailbox with its move type (`Primary`, `PrimaryAndArchive`, `PrimaryOnly`, `ArchiveOnly`), target database(s) and batch. Targets by `DatabaseMap`, then the least loaded target (existing data counted); archives by map, dedicated archive databases (`ArchiveTargetDatabasePattern`) or with their primary. Batches `Balanced` (equal volume, counts differing by one at most), `PerSourceDatabase` or `PerTargetDatabase`. Saved as `MigrationPlan.csv` + `.json`; Start uses the latest plan of the workload (or `-PlanPath`) and refuses a plan older than `Plan.MaxPlanAgeDays` without `-Force`.
- **Pre-flight** at Start: deleted or already moved mailboxes, part already moved (move type reduced), dismounted or no longer valid targets, moves in progress, users already in a batch, batch names already used (finished batch of the tool replaced, active or foreign batch skipped), plan made with another naming refused. Confirmation question with the totals; `-WhatIf` sends every Exchange command with `-WhatIf`.
- **Status** from the move request statistics joined with the migration users: per batch and per mailbox, progress, size transferred, stalled moves (`StatusDetail`), failures with their message, next command. `-Follow` refreshes it (console line + auto-reloading HTML) until the batches are finished.
- **Completion** now (`Complete-MigrationBatch` + public folder moves resumed) or scheduled (`-CompleteAfter`: `Set-MoveRequest -CompleteAfter` + `Resume-MoveRequest` per move, the method measured to work on Exchange 2019); `-Follow` after completion.
- **Cleanup** limited to the objects of the tool, computed and shown before confirmation.
- **Console** in the style of the author's tools: title card, numbered steps with icons, tables, framed summary; emoji in Windows Terminal, console-font symbols elsewhere (`EMM_ICONS`, `NO_COLOR`, `EMM_FORCE_COLOR`); progress bars for long reads and waits.
- **HTML report** for every execution (`templates\Report.template.html`, same design system as Purview DLP Report: light/dark theme, tiles, distribution bars, batch cards, database load bars, *Next step* box with the command to copy, filterable and sortable grid with details and CSV export, auto-refresh in follow mode) and **CSV** files (UTF-8 BOM, French decimal comma with `;`).
- **Daily log** with every step, every result and every change command with its parameters (`[CHANGE]`, `[WhatIf]`).
- **Tests**: `tests\FakeExchange.ps1` (fictitious organisation with the Exchange cmdlets in memory, refusing what Exchange refuses), 44 Pester tests, and `tests\Invoke-EndToEnd.ps1` (47 checks of the whole lifecycle with the real script, run in Windows PowerShell 5.1 and PowerShell 7).
- **Tools**: `tools\New-EmmPackage.ps1` (files needed to run, optional organisation configuration, checked), `tools\Build-Documentation.ps1` (guide in HTML, hero badges from the front matter).
- **Administrator guide** (`docs\ExchangeMailboxMigration-Guide.md` / `.html`) with screenshots of a fictitious organisation, recipes, troubleshooting and the mapping from version 1.

### Fixed (compared with Migrate-ExchangeBals 1.x)
- Step 20 removed **every move request of the organisation** and every `BatchNN` batch before submitting, including moves in progress of other tools. Now: decisions per mailbox and per batch; finished move requests are removed only when they belong to the tool (`Move.ReplaceFinishedMoveRequests = 'Tool' | 'All' | 'None'`); active batches and moves in progress are never touched.
- Step 22 removed **Synced** batches whose moves were not completed (synchronised moves lost). Now a batch is removed only when all its moves (public folders included) are finished; batches with failed moves are kept unless `-IncludeFailed`; orphan migration users are limited to the tool's batches (and to `-Batch`), and the check fails closed.
- Public folder mailboxes were moved with `New-MoveRequest -PublicFolder`, a parameter that does not exist, and never resumed. Now a plain suspended move request labelled with the batch, resumed by Complete (also after the batch is completed, and for batches made only of public folder mailboxes).
- Archive-only moves were sent without `MailboxType`: Exchange moved the primary too, to a database of its choice. Move types and targets are now always explicit.
- Plan and execution disagreed on archive targets (map ignored at submission); primary and archive were balanced separately then regrouped. One row per mailbox, one target choice, used by Start.
- System moves used `BadItemLimit 100` without `AcceptLargeDataLoss` (refused by Exchange). Limits come from the configuration and the Exchange rules are checked; `LargeItemLimit` is no longer sent to `New-MigrationBatch -Local`, which does not support it.
- Follow-up and scheduled completion were reachable only through environment variables, and the HTML suggested commands of another script. Real parameters now, and the right command in every report.
- The database map was guessed from the digits of the patterns. Explicit `DatabaseMap`, otherwise balancing.
- Sizes ignored recoverable items. Sizes are now items + recoverable items (what a move copies), read with one `Get-MailboxStatistics -Database` per source database.

### Documented
- PowerShell pitfalls met during the build (guide, chapter 13): `@()` on a list created with `New-Object` throws; `.Count` under StrictMode on single objects; `$'` in regex replacements; `-WhatIf` on a script affects file writing.
- Monitoring mailboxes left on the old databases and how to retire them (guide, Annex A).
