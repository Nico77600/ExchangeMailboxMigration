#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Mailbox Migration - PowerShell module.

.DESCRIPTION
    Helper functions used by Invoke-ExchangeMailboxMigration.ps1. The module is organised in
    regions, in the order of an execution:

        1. Console and log        Write-Emm* functions (what the administrator sees, also logged)
        2. Configuration          Import-EmmConfiguration, Resolve-EmmScope
        3. Exchange connection    Connect-EmmExchange / Disconnect-EmmExchange
        4. Inventory              databases, mailboxes, existing migration objects
        5. Plan                   New-EmmPlan, Save-EmmPlan, Import-EmmPlan, Find-EmmPlan
        6. Migration actions      Start-EmmMigration, Complete-EmmBatch, Invoke-EmmCleanup
        7. Status                 Get-EmmStatus
        8. Reports                New-EmmReport (templates\Report.template.html + CSV)

    Every value read from Exchange goes through a small adapter (Get-EmmProp, ConvertTo-EmmMB,
    ConvertTo-EmmMailboxRecord...): the rest of the module only handles its own objects, whose
    properties are known. This keeps the code identical with remote PowerShell (deserialized
    objects, sizes as text) and with the local snap-in (live objects).

    Windows PowerShell 5.1 only (Exchange Management Shell): PowerShell 7 is not supported by Microsoft
    for Exchange Server management, and the module refuses to load in it.
    The code is ASCII only: icons and frame characters are built from their code points.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    History : see CHANGELOG.md
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    throw "Exchange Mailbox Migration runs in Windows PowerShell 5.1 only (Exchange Management Shell). PowerShell $($PSVersionTable.PSVersion) is not supported by Microsoft for Exchange Server management."
}

$script:ToolName = 'Exchange Mailbox Migration'
$script:ToolVersion = '2.0.0'
$script:ToolRoot = $PSScriptRoot
$script:LogWriter = $null
$script:LogPath = $null
$script:Actions = [System.Collections.Generic.List[object]]::new()
$script:ExchangeSession = $null
$script:ExchangeModule = $null
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture

# Mailbox types known by the tool. Monitoring mailboxes are never part of any list: they are
# never moved (they belong to the Managed Availability of each server and are recreated by it).
$script:SystemTypes = @('ArbitrationMailbox', 'AuditLogMailbox', 'AuxAuditLogMailbox', 'DiscoveryMailbox')
$script:UserTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'LinkedMailbox', 'LinkedRoomMailbox',
    'TeamMailbox', 'PublicFolderMailbox', 'Archive')
# Move request statuses after which a request no longer moves anything.
$script:FinishedMoveStatuses = @('Completed', 'CompletedWithWarning', 'Failed')
# Migration batch statuses after which a batch no longer moves anything.
$script:FinishedBatchStatuses = @('Completed', 'CompletedWithErrors', 'Failed', 'Stopped', 'Corrupted')

# ---------------------------------------------------------------------------------------------
# Console theme: colours and icons.
#   - Colours (ANSI) are used when the console supports them and the output is not redirected
#     (scheduled task, log capture). NO_COLOR disables them, EMM_FORCE_COLOR=1 forces them.
#   - Icons: emoji in modern terminals (Windows Terminal, VS Code), simple symbols elsewhere.
#     Emoji are chosen among those always two columns wide. The symbols used outside modern
#     terminals all exist in the classic console fonts (code page 437 or Latin-1), so the
#     Exchange Management Shell shows them. Force a style with EMM_ICONS = Emoji | Symbols | Ascii.
# ---------------------------------------------------------------------------------------------
function Test-EmmAnsi {
    if ($env:EMM_FORCE_COLOR -eq '1') { return $true }
    if ($env:NO_COLOR) { return $false }
    if ([Console]::IsOutputRedirected) { return $false }
    try { return [bool]$Host.UI.SupportsVirtualTerminal } catch { return $false }
}

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Cyan = ''; Green = ''; Yellow = ''; Red = ''; White = ''; Blue = ''; Violet = '' }
if (Test-EmmAnsi) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Cyan = "$e[38;2;97;214;214m"; Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"
        Blue = "$e[38;2;96;165;250m"; Violet = "$e[38;2;167;139;250m"
    }
}
$script:IconStyle = if ($env:EMM_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:EMM_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

function Get-EmmIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4EC; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C
                Info = & $u 0x1F539; Skip = & $u 0x23E9; Database = & $u 0x1F4BE; Plan = & $u 0x1F4CB; Key = & $u 0x1F510
                Mailbox = & $u 0x1F4EB; System = & $u 0x1F527; User = & $u 0x1F464; People = & $u 0x1F465; Archive = & $u 0x1F4E6
                Batch = & $u 0x1F4DA; Move = & $u 0x1F69A; Report = & $u 0x1F4CA; File = & $u 0x1F4C4; Folder = & $u 0x1F4C1
                Clock = & $u 0x23F3; Log = & $u 0x1F4DD; Done = & $u 0x1F389; Chart = & $u 0x1F4C8; Calendar = & $u 0x1F4C5
                Target = & $u 0x1F3AF; Broom = & $u 0x1F9F9; Rocket = & $u 0x1F680; Search = & $u 0x1F50E; Stop = & $u 0x1F6D1
                Sync = & $u 0x1F504; Flag = & $u 0x1F3C1; Lock = & $u 0x1F512; Simulate = & $u 0x1F9EA; Next = & $u 0x1F449
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Database = & $u 0x25A0; Plan = & $u 0x25BA; Key = & $u 0x2194; Mailbox = '@'
                System = & $u 0x2666; User = & $u 0x263A; People = & $u 0x263B; Archive = & $u 0x00A7; Batch = & $u 0x2261
                Move = & $u 0x2192; Report = & $u 0x2261; File = & $u 0x25AC; Folder = & $u 0x2302; Clock = & $u 0x25CB
                Log = & $u 0x00B6; Done = & $u 0x221A; Chart = & $u 0x2191; Calendar = & $u 0x263C; Target = & $u 0x25D9
                Broom = & $u 0x00A4; Rocket = & $u 0x2191; Search = & $u 0x25BA; Stop = & $u 0x25A0; Sync = & $u 0x2195
                Flag = & $u 0x25BC; Lock = & $u 0x2666; Simulate = & $u 0x2248; Next = & $u 0x25BA
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Database = '#'; Plan = '?'; Key = '@'
                Mailbox = '@'; System = '%'; User = 'u'; People = '&'; Archive = 'a'; Batch = '='; Move = '>'; Report = '='
                File = '-'; Folder = '>'; Clock = '~'; Log = '='; Done = '*'; Chart = '^'; Calendar = ':'; Target = 'o'
                Broom = '~'; Rocket = '^'; Search = '?'; Stop = '#'; Sync = '~'; Flag = 'v'; Lock = '%'; Simulate = '~'; Next = '>'
            }
        }
    }
}

function Get-EmmFrame {
    <#
    .SYNOPSIS
        Frame characters: rounded corners in modern terminals (emoji style), square corners in the
        classic console (present in every console font), plain ASCII with the Ascii style.
    #>
    switch ($script:IconStyle) {
        'Emoji' { return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 } }
        'Symbols' { return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 } }
        default { return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' } }
    }
}

$script:Icons = Get-EmmIconSet $script:IconStyle
# Emoji are two columns wide in the console; symbols are one: pad symbols so text stays aligned.
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$script:Dot = [char]0x00B7
$script:Arrow = if ($script:IconStyle -eq 'Ascii') { '->' } else { [string][char]0x2192 }

#region 1. Console and log ---------------------------------------------------------------

function Get-EmmIcon {
    param([Parameter(Mandatory)][string]$Name)
    $icon = $script:Icons[$Name]
    if (-not $icon) { $icon = $script:Icons['Info'] }
    return $icon + $script:IconPad
}

function Format-EmmNumber {
    param([Parameter(Mandatory)][AllowNull()]$Value)
    if ($null -eq $Value -or "$Value" -eq '') { return '-' }
    return ([double]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Format-EmmDuration {
    param([Parameter(Mandatory)][double]$Seconds)
    # 0.0 (not 0): with an integer first argument PowerShell picks Math.Max(int, int) and drops the decimals.
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalDays -ge 2) { return '{0} d {1:00} h' -f [int][Math]::Floor($t.TotalDays), $t.Hours }
    if ($t.TotalHours -ge 1) { return '{0} h {1:00} min' -f [int][Math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return '{0} min {1:00} s' -f $t.Minutes, $t.Seconds }
    return ('{0:0.0} s' -f $t.TotalSeconds)
}

function Format-EmmSize {
    <# Size given in MB, shown in MB, GB or TB. #>
    param([Parameter(Mandatory)][AllowNull()]$MB)
    if ($null -eq $MB -or "$MB" -eq '') { return '-' }
    $v = [double]$MB
    $culture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    if ($v -ge 1048576) { return ($v / 1048576).ToString('0.00', $culture) + ' TB' }
    if ($v -ge 1024) { return ($v / 1024).ToString('0.0', $culture) + ' GB' }
    return $v.ToString('0', $culture) + ' MB'
}

function Format-EmmDate {
    <# Date as 'yyyy-MM-dd HH:mm:ss' (local time of the server running the tool), '' when empty or unreadable. #>
    param([AllowNull()]$Value)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant) }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$d)) { return $d.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant) }
    return [string]$Value
}

function Get-EmmBar {
    <# Text progress bar, e.g. "#########-----" with the block characters of the console fonts. #>
    param([double]$Percent, [int]$Width = 20)
    $full = if ($script:IconStyle -eq 'Ascii') { '#' } else { [string][char]0x2588 }
    $empty = if ($script:IconStyle -eq 'Ascii') { '.' } else { [string][char]0x2591 }
    $n = [int][Math]::Round([Math]::Max(0.0, [Math]::Min(100.0, $Percent)) * $Width / 100)
    return ($full * $n) + ($empty * ($Width - $n))
}

function Start-EmmLog {
    <# Opens (or continues) today's log file and deletes log files older than the retention. #>
    param([Parameter(Mandatory)][string]$Directory, [int]$RetentionDays = 90)
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('ExchangeMailboxMigration_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = New-Object IO.FileStream($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:LogWriter = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
    $script:LogWriter.AutoFlush = $true
    $limit = (Get-Date).AddDays(-$RetentionDays)
    foreach ($old in [IO.Directory]::GetFiles($Directory, 'ExchangeMailboxMigration_*.log')) {
        if ([IO.File]::GetLastWriteTime($old) -lt $limit) { try { [IO.File]::Delete($old) } catch { } }
    }
    return $script:LogPath
}

function Stop-EmmLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-EmmLog {
    <# Writes one line to the log file only (never to the console). The log never contains colours or icons. #>
    param([ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'CHANGE', 'DEBUG')][string]$Level = 'INFO', [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if ($script:LogWriter) {
        $script:LogWriter.WriteLine(('{0} [{1,-6}] {2}' -f (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffzzz', $script:Invariant), $Level, $Message))
    }
}

function Write-EmmBanner {
    <#
    .SYNOPSIS
        Title card at the start of an execution:

          +----------------------------------------------------------------------------+
          |  @  Exchange Mailbox Migration                      v2.0.0 . Nicolas Fabert |
          |     Plan . user mailboxes                                                  |
          +----------------------------------------------------------------------------+
             >  Mode       Plan
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle, [System.Collections.Specialized.OrderedDictionary]$Details)
    $C = $script:C; $F = Get-EmmFrame; $width = 76
    $right = "v$($script:ToolVersion) $($script:Dot) Nicolas Fabert"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $iconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, ([string]$F.Horizontal * $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, (' ' * $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        $sub = "     $Subtitle"
        if ($sub.Length -gt $width) { $sub = $sub.Substring(0, $width) }
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, $sub.PadRight($width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, ([string]$F.Horizontal * $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            if ($value -is [array]) { $icon = Get-EmmIcon $value[0]; $text = $value[1] } else { $icon = '   '; $text = $value }
            Write-Host ('     {0}{1}{2,-10}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
    Write-EmmLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) { foreach ($key in $Details.Keys) { $v = $Details[$key]; Write-EmmLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v })) } }
}

function Write-EmmStep {
    <#
    .SYNOPSIS
        Step header with a coloured number pill and an icon, e.g.

           3/5   [key]  Connecting to Exchange
    #>
    param([Parameter(Mandatory)][int]$Number, [Parameter(Mandatory)][int]$Total, [Parameter(Mandatory)][string]$Title, [string]$Icon = 'Info')
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-EmmIcon $Icon), $C.Bold, $Title)
    Write-EmmLog 'STEP' "[$Number/$Total] $Title"
}

function Write-EmmItem {
    <# One indented result line with a status icon, also written to the log. #>
    param([ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip')][string]$Status = 'Info', [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Icon)
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    $symbol = Get-EmmIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
    Write-EmmLog $level $Text
}

function Write-EmmTable {
    <#
    .SYNOPSIS
        Aligned table under a step (one row per batch or per database), also written to the log.
    .PARAMETER Columns
        Ordered list: key = column title, value = width (negative = left aligned, positive = right aligned).
    .PARAMETER Rows
        Objects with one property per column title, plus an optional 'Status' (Ok | Warn | Fail | Info | Skip).
    #>
    param([Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Columns, [AllowEmptyCollection()][object[]]$Rows = @())
    $C = $script:C
    $head = foreach ($k in $Columns.Keys) { "{0,$($Columns[$k])}" -f $k }
    Write-Host ('      {0}{1}{2}{3}' -f $C.Dim, ('  ' + $script:IconPad), ($head -join '  '), $C.Reset)
    foreach ($row in $Rows) {
        $status = if ($row.PSObject.Properties['Status']) { [string]$row.Status } else { 'Info' }
        $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red; Info = $C.Dim; Skip = $C.Dim }[$status]
        if (-not $color) { $color = $C.Dim }
        $icon = if ($status -in 'Ok', 'Warn', 'Fail', 'Skip') { Get-EmmIcon $status } else { Get-EmmIcon 'Info' }
        $cells = foreach ($k in $Columns.Keys) {
            $w = [int]$Columns[$k]; $text = [string]$row.$k
            if ($text.Length -gt [Math]::Abs($w)) { $text = $text.Substring(0, [Math]::Abs($w) - 1) + '~' }
            "{0,$w}" -f $text
        }
        Write-Host ('      {0}{1}{2}{3}' -f $color, $icon, $C.Reset, ($cells -join '  '))
        Write-EmmLog $(if ($status -eq 'Fail') { 'ERROR' } elseif ($status -eq 'Warn') { 'WARN' } else { 'INFO' }) ((@(foreach ($k in $Columns.Keys) { '{0}={1}' -f $k, $row.$k }) -join ' | '))
    }
}

function Write-EmmSummary {
    <#
    .SYNOPSIS
        Final summary card:

          +- [done]  Plan ready ---------------------------------------------------+
            [mailbox]  Mailboxes   ...
          +------------------------------------------------------------------------+
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok')
    $C = $script:C; $F = Get-EmmFrame; $width = 76
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), (([string]$F.Horizontal * $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        if ($value -is [array]) { $rowIcon = Get-EmmIcon $value[0]; $text = $value[1] } else { $rowIcon = '   '; $text = $value }
        Write-Host ('    {0}{1}{2,-11}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
        Write-EmmLog 'INFO' ('Summary - {0}: {1}' -f $key, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, ([string]$F.Horizontal * $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

#endregion
#region 2. Configuration ------------------------------------------------------------------

function Resolve-EmmPath {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if (-not [IO.Path]::IsPathRooted($expanded)) { $expanded = Join-Path $Root $expanded }
    return [IO.Path]::GetFullPath($expanded)
}

function Import-EmmConfiguration {
    <#
    .SYNOPSIS
        Reads the configuration file, checks every value and returns it with absolute paths.
        All problems are reported together so the administrator can fix them in one go.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Root = $script:ToolRoot)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    try { $config = Import-PowerShellDataFile -LiteralPath $Path }
    catch { throw "The configuration file is not valid PowerShell data ($Path): $($_.Exception.Message)" }

    $errors = [System.Collections.Generic.List[string]]::new()
    $sections = 'Connection', 'Databases', 'Scope', 'Plan', 'Move', 'System', 'Status', 'Cleanup', 'Report', 'Logging'
    foreach ($section in $sections) {
        if (-not $config.ContainsKey($section) -or $config[$section] -isnot [hashtable]) { $errors.Add("Section '$section' is missing.") }
    }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }

    function Get-Value([hashtable]$Section, [string]$Key, $Default) {
        if ($Section.ContainsKey($Key) -and $null -ne $Section[$Key]) { return $Section[$Key] }
        return $Default
    }
    function Test-Int($Value, [string]$Name, [long]$Min, [long]$Max) {
        $n = 0L
        if (-not [long]::TryParse("$Value", [ref]$n) -or $n -lt $Min -or $n -gt $Max) { $errors.Add("$Name must be a whole number between $Min and $Max (current value: '$Value').") }
        return [int]$n
    }
    function Test-Bool($Value, [string]$Name) {
        if ($Value -isnot [bool]) { $errors.Add("$Name must be `$true or `$false (current value: '$Value')."); return $false }
        return $Value
    }
    function Test-Choice($Value, [string]$Name, [string[]]$Allowed) {
        if ("$Value" -notin $Allowed) { $errors.Add("$Name must be $($Allowed -join ' | ') (current value: '$Value').") }
        return [string]$Value
    }
    function Test-Regex($Value, [string]$Name) {
        if ("$Value") { try { [void][regex]::new("$Value") } catch { $errors.Add("$Name is not a valid regular expression: $($_.Exception.Message)") } }
        return [string]$Value
    }
    function Get-List($Value) { return @(@($Value) | Where-Object { $null -ne $_ -and "$_" -ne '' } | ForEach-Object { "$_".Trim() }) }

    $cn = $config.Connection; $db = $config.Databases; $sc = $config.Scope; $pl = $config.Plan; $mv = $config.Move
    $sy = $config.System; $st = $config.Status; $cl = $config.Cleanup; $rp = $config.Report; $lg = $config.Logging

    $map = [ordered]@{}
    $rawMap = Get-Value $db 'DatabaseMap' @{}
    if ($rawMap -isnot [hashtable]) { $errors.Add('Databases.DatabaseMap must be a table: @{ ''DB01'' = ''DB-01'' }.') }
    else { foreach ($k in @($rawMap.Keys | Sort-Object)) { if ("$k" -and "$($rawMap[$k])") { $map["$k".Trim()] = "$($rawMap[$k])".Trim() } } }

    $settings = [ordered]@{
        ConfigPath = [IO.Path]::GetFullPath($Path)
        Connection = [ordered]@{
            ExchangeServer   = [string](Get-Value $cn 'ExchangeServer' '')
            Authentication   = Test-Choice (Get-Value $cn 'Authentication' 'Kerberos') 'Connection.Authentication' @('Kerberos', 'Negotiate', 'Default')
            DomainController = [string](Get-Value $cn 'DomainController' '')
        }
        Databases = [ordered]@{
            SourceDatabasePattern        = Test-Regex (Get-Value $db 'SourceDatabasePattern' '') 'Databases.SourceDatabasePattern'
            TargetDatabasePattern        = Test-Regex (Get-Value $db 'TargetDatabasePattern' '') 'Databases.TargetDatabasePattern'
            ArchiveTargetDatabasePattern = Test-Regex (Get-Value $db 'ArchiveTargetDatabasePattern' '') 'Databases.ArchiveTargetDatabasePattern'
            DatabaseMap                  = $map
            CountExistingData            = Test-Bool (Get-Value $db 'CountExistingData' $true) 'Databases.CountExistingData'
        }
        Scope = [ordered]@{
            Workload           = Test-Choice (Get-Value $sc 'Workload' 'All') 'Scope.Workload' @('System', 'User', 'All')
            # @( ) around every list: a function returning one item (or none) gives a scalar (or $null).
            SystemMailboxTypes = @(Get-List (Get-Value $sc 'SystemMailboxTypes' $script:SystemTypes))
            UserMailboxTypes   = @(Get-List (Get-Value $sc 'UserMailboxTypes' @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'PublicFolderMailbox', 'Archive')))
            ExcludeMailboxes   = @(Get-List (Get-Value $sc 'ExcludeMailboxes' @()))
        }
        Plan = [ordered]@{
            BatchCount      = Test-Int (Get-Value $pl 'BatchCount' 12) 'Plan.BatchCount' 1 99
            BatchStrategy   = Test-Choice (Get-Value $pl 'BatchStrategy' 'Balanced') 'Plan.BatchStrategy' @('Balanced', 'PerSourceDatabase', 'PerTargetDatabase')
            BatchNamePrefix = [string](Get-Value $pl 'BatchNamePrefix' 'Batch')
            MaxPlanAgeDays  = Test-Int (Get-Value $pl 'MaxPlanAgeDays' 7) 'Plan.MaxPlanAgeDays' 1 365
        }
        Move = [ordered]@{
            BadItemLimit                = Test-Int (Get-Value $mv 'BadItemLimit' 20) 'Move.BadItemLimit' 0 100000
            LargeItemLimit              = Test-Int (Get-Value $mv 'LargeItemLimit' 0) 'Move.LargeItemLimit' 0 100000
            AcceptLargeDataLoss         = Test-Bool (Get-Value $mv 'AcceptLargeDataLoss' $false) 'Move.AcceptLargeDataLoss'
            NotificationEmails          = @(Get-List (Get-Value $mv 'NotificationEmails' @()))
            ReplaceFinishedMoveRequests = Test-Choice $(switch ((Get-Value $mv 'ReplaceFinishedMoveRequests' 'Tool')) { $true { 'Tool' } $false { 'None' } default { $_ } }) 'Move.ReplaceFinishedMoveRequests' @('Tool', 'All', 'None')
        }
        System = [ordered]@{
            BatchName          = [string](Get-Value $sy 'BatchName' 'EMM-SystemMailboxes')
            TargetDatabase     = [string](Get-Value $sy 'TargetDatabase' '')
            AllowLargeItems    = Test-Bool (Get-Value $sy 'AllowLargeItems' $true) 'System.AllowLargeItems'
            WaitForCompletion  = Test-Bool (Get-Value $sy 'WaitForCompletion' $true) 'System.WaitForCompletion'
            WaitTimeoutMinutes = Test-Int (Get-Value $sy 'WaitTimeoutMinutes' 30) 'System.WaitTimeoutMinutes' 1 1440
            PollSeconds        = Test-Int (Get-Value $sy 'PollSeconds' 30) 'System.PollSeconds' 1 600
        }
        Status = [ordered]@{
            RefreshMinutes = Test-Int (Get-Value $st 'RefreshMinutes' 2) 'Status.RefreshMinutes' 1 1440
            OpenReport     = Test-Bool (Get-Value $st 'OpenReport' $true) 'Status.OpenReport'
        }
        Cleanup = [ordered]@{
            IncludeFailed              = Test-Bool (Get-Value $cl 'IncludeFailed' $false) 'Cleanup.IncludeFailed'
            RemoveOrphanMigrationUsers = Test-Bool (Get-Value $cl 'RemoveOrphanMigrationUsers' $true) 'Cleanup.RemoveOrphanMigrationUsers'
            SettleSeconds              = Test-Int (Get-Value $cl 'SettleSeconds' 15) 'Cleanup.SettleSeconds' 0 300
        }
        Report = [ordered]@{
            OutputPath   = Resolve-EmmPath ([string](Get-Value $rp 'OutputPath' '.\reports')) $Root
            CsvDelimiter = [string](Get-Value $rp 'CsvDelimiter' ';')
            Organization = [string](Get-Value $rp 'Organization' '')
            TemplatePath = Join-Path $Root 'templates\Report.template.html'
        }
        Logging = [ordered]@{
            Path          = Resolve-EmmPath ([string](Get-Value $lg 'Path' '.\logs')) $Root
            RetentionDays = Test-Int (Get-Value $lg 'RetentionDays' 90) 'Logging.RetentionDays' 1 3650
        }
    }

    # ---- Cross checks ------------------------------------------------------------------------
    $d = $settings.Databases
    if (-not $d.SourceDatabasePattern -and $d.DatabaseMap.Count -eq 0) { $errors.Add('Databases: set SourceDatabasePattern or DatabaseMap (the databases to empty).') }
    if (-not $d.TargetDatabasePattern -and $d.DatabaseMap.Count -eq 0) { $errors.Add('Databases: set TargetDatabasePattern or DatabaseMap (the databases to fill).') }
    foreach ($k in $d.DatabaseMap.Keys) { if ($k -eq $d.DatabaseMap[$k]) { $errors.Add("Databases.DatabaseMap: '$k' cannot be mapped to itself.") } }
    $s = $settings.Scope
    foreach ($t in @($s.SystemMailboxTypes) + @($s.UserMailboxTypes)) {
        if ($t -eq 'MonitoringMailbox') { $errors.Add('Monitoring mailboxes are never moved: remove MonitoringMailbox from the Scope section.') }
    }
    foreach ($t in $s.SystemMailboxTypes) { if ($t -ne 'MonitoringMailbox' -and $t -notin $script:SystemTypes) { $errors.Add("Scope.SystemMailboxTypes: '$t' is not a system mailbox type ($($script:SystemTypes -join ', ')).") } }
    foreach ($t in $s.UserMailboxTypes) { if ($t -ne 'MonitoringMailbox' -and $t -notin $script:UserTypes) { $errors.Add("Scope.UserMailboxTypes: '$t' is not a user mailbox type ($($script:UserTypes -join ', ')).") } }
    $m = $settings.Move
    if (($m.BadItemLimit -ge 51 -or $m.LargeItemLimit -ge 51) -and -not $m.AcceptLargeDataLoss) {
        $errors.Add('Move: BadItemLimit or LargeItemLimit is 51 or more: Exchange requires Move.AcceptLargeDataLoss = $true (accept to lose these items).')
    }
    if ($settings.Plan.BatchNamePrefix -notmatch '^[A-Za-z][A-Za-z0-9_-]{0,40}$') { $errors.Add('Plan.BatchNamePrefix must start with a letter and contain only letters, digits, - or _ (40 characters at most).') }
    if ($settings.System.BatchName -notmatch '^[A-Za-z][A-Za-z0-9_-]{0,60}$') { $errors.Add('System.BatchName must start with a letter and contain only letters, digits, - or _.') }
    if ($settings.System.BatchName -match ('^' + [regex]::Escape($settings.Plan.BatchNamePrefix) + '\d+$')) { $errors.Add('System.BatchName must not look like a user batch name (Plan.BatchNamePrefix + number).') }
    if ($settings.Report.CsvDelimiter -notin ';', ',', "`t", '|') { $errors.Add("Report.CsvDelimiter must be ';', ',', '|' or a tab.") }
    foreach ($a in $settings.Move.NotificationEmails) { if ($a -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $errors.Add("Move.NotificationEmails: '$a' is not an e-mail address.") } }
    if (-not (Test-Path -LiteralPath $settings.Report.TemplatePath -PathType Leaf)) { $errors.Add("Report template not found: $($settings.Report.TemplatePath)") }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }
    return $settings
}

function Resolve-EmmScope {
    <#
    .SYNOPSIS
        Turns the configuration and the command line (-Workload, -MailboxType) into the effective
        selection of this execution.
    .DESCRIPTION
        Without -MailboxType: the workload (command line, otherwise Scope.Workload) selects the
        system types and/or the user types of the configuration.
        With -MailboxType: exactly these types (any known type, even if absent from the
        configuration lists). The workload is deduced from them when -Workload is not given.
        'Archive' selects the archive mailboxes: of the user types also selected, or of every
        user type when 'Archive' is alone.
    #>
    param([Parameter(Mandatory)]$Settings, [string]$Workload, [string[]]$MailboxType)
    $types = @($MailboxType | Where-Object { $_ } | ForEach-Object { ($_ -split ',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($t in $types) {
        if ($t -eq 'MonitoringMailbox') { throw 'Monitoring mailboxes are never moved (MonitoringMailbox cannot be selected).' }
        if ($t -notin ($script:SystemTypes + $script:UserTypes)) { throw "Unknown mailbox type '$t'. Known types: $(($script:SystemTypes + $script:UserTypes) -join ', ')." }
    }
    if ($types.Count) {
        $hasSystem = @($types | Where-Object { $_ -in $script:SystemTypes }).Count -gt 0
        $hasUser = @($types | Where-Object { $_ -in $script:UserTypes }).Count -gt 0
        if (-not $Workload) { $Workload = if ($hasSystem -and $hasUser) { 'All' } elseif ($hasSystem) { 'System' } else { 'User' } }
        $system = @($types | Where-Object { $_ -in $script:SystemTypes -and $Workload -in 'System', 'All' })
        $user = @($types | Where-Object { $_ -in $script:UserTypes -and $Workload -in 'User', 'All' })
        if (-not ($system.Count + $user.Count)) { throw "-MailboxType $($types -join ',') selects nothing in the workload '$Workload'." }
    } else {
        if (-not $Workload) { $Workload = $Settings.Scope.Workload }
        $system = @(if ($Workload -in 'System', 'All') { $Settings.Scope.SystemMailboxTypes })
        $user = @(if ($Workload -in 'User', 'All') { $Settings.Scope.UserMailboxTypes })
    }
    $primary = @($user | Where-Object { $_ -ne 'Archive' })
    $includeArchives = $user -contains 'Archive'
    $owners = @($primary | Where-Object { $_ -ne 'PublicFolderMailbox' })
    if ($includeArchives -and -not $owners.Count) { $owners = @($script:UserTypes | Where-Object { $_ -notin 'Archive', 'PublicFolderMailbox' }) }
    $label = switch ($Workload) { 'System' { 'system mailboxes' } 'User' { 'user mailboxes' } default { 'system and user mailboxes' } }
    if ($types.Count) { $label += ' (' + ($types -join ', ') + ')' }
    return [pscustomobject]@{
        Workload           = $Workload
        SystemTypes        = $system
        UserPrimaryTypes   = $primary
        IncludeArchives    = $includeArchives
        ArchiveOwnerTypes  = @(if ($includeArchives) { $owners })
        ExplicitTypes      = $types
        Label              = $label
    }
}

#endregion
#region 3. Exchange connection ------------------------------------------------------------------

$script:ExchangeCommands = @(
    'Get-ExchangeServer', 'Get-MailboxDatabase', 'Get-Mailbox', 'Get-MailboxStatistics', 'Get-Recipient',
    'Get-MoveRequest', 'Get-MoveRequestStatistics', 'New-MoveRequest', 'Set-MoveRequest', 'Resume-MoveRequest', 'Remove-MoveRequest',
    'Get-MigrationBatch', 'New-MigrationBatch', 'Start-MigrationBatch', 'Complete-MigrationBatch', 'Remove-MigrationBatch',
    'Get-MigrationUser', 'Get-MigrationUserStatistics', 'Remove-MigrationUser')

function Test-EmmExchangeLoaded {
    return [bool](Get-Command Get-MailboxDatabase -ErrorAction SilentlyContinue) -and [bool](Get-Command Get-MigrationBatch -ErrorAction SilentlyContinue)
}

function Connect-EmmExchange {
    <#
    .SYNOPSIS
        Makes the Exchange cmdlets available to the module and checks the account can run them.
    .DESCRIPTION
        In this order:
          1. cmdlets already loaded (Exchange Management Shell, or an existing remote session);
          2. remote PowerShell to Connection.ExchangeServer (or the local server when it is empty);
          3. the local snap-in (on an Exchange server only, when no server is configured).
        Remote PowerShell only exposes the cmdlets allowed by the RBAC roles of the account:
        a missing cmdlet therefore means a missing role, and is reported as such.
    .PARAMETER RequiredCommand
        Cmdlets the current mode needs.
    #>
    param([string]$Server, [pscredential]$Credential, [string]$Authentication = 'Kerberos', [string[]]$RequiredCommand = $script:ExchangeCommands)
    $method = $null
    if (Test-EmmExchangeLoaded) {
        $method = 'cmdlets already loaded'
    } else {
        $computer = if ($Server) { $Server } else { [Environment]::MachineName }
        try {
            $params = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = "http://$computer/PowerShell/"
                Authentication    = $Authentication
                AllowRedirection  = $true
                ErrorAction       = 'Stop'
            }
            if ($Credential) { $params['Credential'] = $Credential }
            Write-EmmLog 'INFO' "New-PSSession http://$computer/PowerShell/ ($Authentication)"
            $script:ExchangeSession = New-PSSession @params
            # Only the cmdlets used by the tool are imported: faster, and nothing else is exposed.
            $script:ExchangeModule = Import-PSSession -Session $script:ExchangeSession -CommandName $script:ExchangeCommands -DisableNameChecking -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue
            $method = "remote PowerShell to $computer"
        } catch {
            $remoteError = $_.Exception.Message
            if ($script:ExchangeSession) { Remove-PSSession $script:ExchangeSession -ErrorAction SilentlyContinue; $script:ExchangeSession = $null }
            if ($Server) { throw "Cannot open remote PowerShell to $computer`: $remoteError" }
            try {
                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
                $method = 'local snap-in'
            } catch {
                throw "Exchange cmdlets not available. Remote PowerShell to $computer failed ($remoteError) and the local snap-in cannot be loaded. Run the tool in the Exchange Management Shell, or set Connection.ExchangeServer."
            }
        }
    }
    $missing = @($RequiredCommand | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count) {
        throw ("Exchange cmdlets not available to this account: {0}. Remote PowerShell only exposes the cmdlets of the RBAC roles of the account: the tool needs the roles Mail Recipients, Move Mailboxes, Migration and View-Only Configuration (for example: Organization Management, or Recipient Management + a role group with Move Mailboxes and Migration)." -f ($missing -join ', '))
    }
    $account = if ($Credential) { $Credential.UserName } else { [Security.Principal.WindowsIdentity]::GetCurrent().Name }
    return [pscustomobject]@{ Method = $method; Account = $account; Server = $(if ($Server) { $Server } else { [Environment]::MachineName }) }
}

function Disconnect-EmmExchange {
    if ($script:ExchangeModule) { Remove-Module $script:ExchangeModule -Force -ErrorAction SilentlyContinue; $script:ExchangeModule = $null }
    if ($script:ExchangeSession) { Remove-PSSession $script:ExchangeSession -ErrorAction SilentlyContinue; $script:ExchangeSession = $null }
}

#endregion
#region 4. Inventory ------------------------------------------------------------------------------------

function Get-EmmProp {
    <# Value of a property, or $Default when the object does not have it (Exchange versions and access modes differ). #>
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

function Get-EmmName {
    <#
    .SYNOPSIS
        Name of a database or server reference: a string with remote PowerShell, an object with a
        Name property with the local snap-in.
    #>
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    $n = Get-EmmProp $Value 'Name'
    if ($n) { return [string]$n }
    return [string]$Value
}

function ConvertTo-EmmMB {
    <#
    .SYNOPSIS
        Size in MB from an Exchange size: '1.5 GB (1,610,612,736 bytes)' (remote PowerShell, any
        regional settings), the same object from the snap-in, a number of bytes, or $null (0).
    #>
    param([AllowNull()]$Size)
    if ($null -eq $Size) { return 0.0 }
    if ($Size -is [int] -or $Size -is [long] -or $Size -is [double] -or $Size -is [decimal] -or $Size -is [uint64]) { return [Math]::Round([double]$Size / 1MB, 2) }
    $text = [string]$Size
    # "(1,610,612,736 bytes)", "(1 610 612 736 bytes)", "(1.610.612.736 octets)"...
    $m = [regex]::Match($text, '\(([0-9][0-9\s,\.''\u00A0\u202F]*)\s*\p{L}+\)')
    if ($m.Success) {
        $digits = $m.Groups[1].Value -replace '[^0-9]', ''
        if ($digits) { return [Math]::Round([double]$digits / 1MB, 2) }
    }
    if ($text -match '^\s*([0-9]+)\s*$') { return [Math]::Round([double]$Matches[1] / 1MB, 2) }
    $value = Get-EmmProp $Size 'Value'
    if ($null -ne $value -and $value.PSObject.Methods['ToBytes']) { return [Math]::Round([double]$value.ToBytes() / 1MB, 2) }
    return 0.0
}

function Get-EmmVersionLabel {
    param([string]$AdminDisplayVersion)
    if ($AdminDisplayVersion -match '15\.2') { return 'Exchange 2019/SE' }
    if ($AdminDisplayVersion -match '15\.1') { return 'Exchange 2016' }
    if ($AdminDisplayVersion -match '15\.0') { return 'Exchange 2013' }
    if ($AdminDisplayVersion -match '14\.') { return 'Exchange 2010' }
    return $AdminDisplayVersion
}

function Get-EmmDatabaseClassification {
    <#
    .SYNOPSIS
        Classifies database names as Source, Target, ArchiveTarget or Other from the Databases
        section of the configuration. No Exchange call: used by the inventory and by the tests.
    .DESCRIPTION
        Source  : key of DatabaseMap, or matches SourceDatabasePattern.
        Target  : matches TargetDatabasePattern, or is a value of DatabaseMap.
        ArchiveTarget : matches ArchiveTargetDatabasePattern (archives only).
        Safety: a database that is a target (of any kind) is never a source.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Name, [Parameter(Mandatory)]$Settings)
    $d = $Settings.Databases
    $warnings = [System.Collections.Generic.List[string]]::new()
    $roles = @{}
    $mapTargets = @($d.DatabaseMap.Values)
    foreach ($n in $Name) {
        $isTarget = [bool](($d.TargetDatabasePattern -and $n -match $d.TargetDatabasePattern) -or ($n -in $mapTargets))
        $isArchiveTarget = [bool]($d.ArchiveTargetDatabasePattern -and $n -match $d.ArchiveTargetDatabasePattern)
        $isSource = [bool]($d.DatabaseMap.Contains($n) -or ($d.SourceDatabasePattern -and $n -match $d.SourceDatabasePattern))
        if ($isSource -and ($isTarget -or $isArchiveTarget)) {
            $warnings.Add("Database $n matches both the source and the target rules: it is treated as a target and is never emptied.")
            $isSource = $false
        }
        $role = if ($isSource) { 'Source' } elseif ($isTarget) { 'Target' } elseif ($isArchiveTarget) { 'ArchiveTarget' } else { 'Other' }
        $roles[$n] = [pscustomobject]@{
            Role = $role; IsSource = $isSource; IsTarget = $isTarget; IsArchiveTarget = $isArchiveTarget
            MapTarget = $(if ($d.DatabaseMap.Contains($n)) { [string]$d.DatabaseMap[$n] } else { '' })
        }
    }
    foreach ($k in @($d.DatabaseMap.Keys)) {
        if ($k -notin $Name) { $warnings.Add("DatabaseMap: the source database $k does not exist.") }
        if ($d.DatabaseMap[$k] -notin $Name) { $warnings.Add("DatabaseMap: the target database $($d.DatabaseMap[$k]) does not exist: the mailboxes of $k will be balanced over the other target databases.") }
    }
    $sources = @($Name | Where-Object { $roles[$_].IsSource } | Sort-Object)
    $targets = @($Name | Where-Object { $roles[$_].IsTarget } | Sort-Object)
    $archiveTargets = @($Name | Where-Object { $roles[$_].IsArchiveTarget } | Sort-Object)
    if (-not $sources.Count) { $warnings.Add('No source database: check Databases.SourceDatabasePattern and Databases.DatabaseMap.') }
    if (-not $targets.Count) { $warnings.Add('No target database: check Databases.TargetDatabasePattern and Databases.DatabaseMap.') }
    if ($d.ArchiveTargetDatabasePattern -and -not $archiveTargets.Count) { $warnings.Add('No database matches Databases.ArchiveTargetDatabasePattern.') }
    return [pscustomobject]@{ Roles = $roles; Sources = $sources; Targets = $targets; ArchiveTargets = $archiveTargets; Warnings = @($warnings) }
}

function Get-EmmDatabaseInventory {
    <# Mailbox databases (recovery databases excluded) with their role, server, version, state and size. #>
    param([Parameter(Mandatory)]$Settings)
    $versions = @{}
    foreach ($s in @(Get-ExchangeServer -ErrorAction Stop -WarningAction SilentlyContinue)) {
        $versions[(Get-EmmName $s.Name)] = Get-EmmVersionLabel ([string](Get-EmmProp $s 'AdminDisplayVersion' ''))
    }
    $raw = @(Get-MailboxDatabase -Status -ErrorAction Stop -WarningAction SilentlyContinue | Where-Object { -not [bool](Get-EmmProp $_ 'Recovery' $false) })
    $class = Get-EmmDatabaseClassification -Name @($raw | ForEach-Object { [string]$_.Name }) -Settings $Settings
    $warnings = [System.Collections.Generic.List[string]]::new()
    foreach ($w in $class.Warnings) { $warnings.Add($w) }
    $list = foreach ($db in $raw) {
        $name = [string]$db.Name
        $server = Get-EmmName (Get-EmmProp $db 'Server' '')
        $r = $class.Roles[$name]
        $mounted = Get-EmmProp $db 'Mounted'
        $item = [pscustomobject][ordered]@{
            Name = $name; Role = $r.Role; IsSource = $r.IsSource; IsTarget = $r.IsTarget; IsArchiveTarget = $r.IsArchiveTarget
            MapTarget = $r.MapTarget; Server = $server; Version = $(if ($versions.ContainsKey($server)) { $versions[$server] } else { '' })
            Mounted = $(if ($null -eq $mounted) { $null } else { [bool]$mounted })
            SizeMB = ConvertTo-EmmMB (Get-EmmProp $db 'DatabaseSize'); WhitespaceMB = ConvertTo-EmmMB (Get-EmmProp $db 'AvailableNewMailboxSpace')
            Mailboxes = 0; Archives = 0; SystemMailboxes = 0; MonitoringMailboxes = 0; ToMove = 0; ToMoveMB = 0.0
            PlannedPrimaries = 0; PlannedArchives = 0; PlannedMB = 0.0
        }
        if (($item.IsTarget -or $item.IsArchiveTarget) -and $item.Mounted -eq $false) { $warnings.Add("Target database $name is not mounted: it receives no mailbox.") }
        $item
    }
    return [pscustomobject]@{ Databases = @($list | Sort-Object Name); Classification = $class; Warnings = @($warnings) }
}

function ConvertTo-EmmMailboxRecord {
    <# Normalized mailbox (the only shape used after reading Exchange). Monitoring mailboxes are recognised here. #>
    param([Parameter(Mandatory)]$Mailbox, [string]$Category)
    $type = [string](Get-EmmProp $Mailbox 'RecipientTypeDetails' '')
    $name = [string](Get-EmmProp $Mailbox 'Name' '')
    if ($type -eq 'MonitoringMailbox' -or $name -like 'HealthMailbox*' -or $Category -eq 'Monitoring') { $Category = 'Monitoring' }
    elseif ($type -in $script:SystemTypes) { $Category = 'System' }
    else { $Category = 'User' }
    $archiveDb = Get-EmmName (Get-EmmProp $Mailbox 'ArchiveDatabase' '')
    $guid = { param($v) $s = [string]$v; if ($s -match '^0{8}-0{4}-0{4}-0{4}-0{12}$') { '' } else { $s.ToLowerInvariant() } }
    return [pscustomobject][ordered]@{
        Name = $name; DisplayName = [string](Get-EmmProp $Mailbox 'DisplayName' $name); Alias = [string](Get-EmmProp $Mailbox 'Alias' '')
        PrimarySmtpAddress = [string](Get-EmmProp $Mailbox 'PrimarySmtpAddress' ''); RecipientTypeDetails = $type; Category = $Category
        Database = Get-EmmName (Get-EmmProp $Mailbox 'Database' ''); ArchiveDatabase = $archiveDb; HasArchive = [bool]$archiveDb
        Server = Get-EmmName (Get-EmmProp $Mailbox 'ServerName' '')
        PrimaryRole = ''; ArchiveRole = ''; Status = ''; Reason = ''; MoveType = ''
        PrimarySizeMB = $null; ArchiveSizeMB = $null; ItemCount = $null; ArchiveItemCount = $null; MoveSizeMB = 0.0
        Guid = & $guid (Get-EmmProp $Mailbox 'Guid' ''); ExchangeGuid = & $guid (Get-EmmProp $Mailbox 'ExchangeGuid' '')
        ArchiveGuid = & $guid (Get-EmmProp $Mailbox 'ArchiveGuid' ''); DistinguishedName = [string](Get-EmmProp $Mailbox 'DistinguishedName' '')
    }
}

function Test-EmmMoveType {
    <# $true when the move type moves the primary (-Part Primary) or the archive (-Part Archive). #>
    param([string]$MoveType, [ValidateSet('Primary', 'Archive')][string]$Part)
    if ($Part -eq 'Primary') { return $MoveType -in 'Primary', 'PrimaryAndArchive', 'PrimaryOnly' }
    return $MoveType -in 'PrimaryAndArchive', 'ArchiveOnly'
}

function Set-EmmMailboxSelection {
    <#
    .SYNOPSIS
        Decides, for every mailbox, whether it is moved in this execution and how (MoveType),
        or why not (Status + Reason). No Exchange call: used by the inventory, the plan and the tests.
    .DESCRIPTION
        Status : ToMove | OnTarget | NotSelected | Outside | Excluded
        MoveType (ToMove only):
          Primary            primary mailbox, no archive
          PrimaryAndArchive  primary and archive
          PrimaryOnly        primary only (the archive stays where it is)
          ArchiveOnly        archive only (the primary is not moved)
        Monitoring mailboxes are always Excluded: this is checked first and cannot be configured.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Mailbox, [Parameter(Mandatory)]$Classification, [Parameter(Mandatory)]$Scope, [Parameter(Mandatory)]$Settings)
    $excluded = @{}
    foreach ($x in $Settings.Scope.ExcludeMailboxes) { $excluded[$x.ToLowerInvariant()] = $true }
    $roleOf = { param($db) if ($db -and $Classification.Roles.ContainsKey($db)) { $Classification.Roles[$db].Role } elseif ($db) { 'Other' } else { '' } }
    $isTargetRole = { param($r) $r -in 'Target', 'ArchiveTarget' }
    foreach ($m in $Mailbox) {
        $m.PrimaryRole = & $roleOf $m.Database
        $m.ArchiveRole = if ($m.HasArchive) { & $roleOf $m.ArchiveDatabase } else { '' }
        $m.MoveType = ''; $m.Reason = ''
        if ($m.Category -eq 'Monitoring') { $m.Status = 'Excluded'; $m.Reason = 'Monitoring mailbox: never moved'; continue }
        $ids = @($m.Name, $m.Alias, $m.PrimarySmtpAddress, $m.Guid, $m.ExchangeGuid) | Where-Object { $_ }
        if (@($ids | Where-Object { $excluded.ContainsKey($_.ToLowerInvariant()) }).Count) { $m.Status = 'Excluded'; $m.Reason = 'Listed in Scope.ExcludeMailboxes'; continue }
        $primaryOnSource = $m.PrimaryRole -eq 'Source'
        $archiveOnSource = $m.HasArchive -and $m.ArchiveRole -eq 'Source'
        if (-not $primaryOnSource -and -not $archiveOnSource) {
            if ((& $isTargetRole $m.PrimaryRole) -and (-not $m.HasArchive -or (& $isTargetRole $m.ArchiveRole))) { $m.Status = 'OnTarget'; $m.Reason = 'Already on a target database' }
            else { $m.Status = 'Outside'; $m.Reason = 'Not on a source database' }
            continue
        }
        if ($m.Category -eq 'System') {
            $primarySelected = $m.RecipientTypeDetails -in $Scope.SystemTypes
            $archiveSelected = $false
        } else {
            $primarySelected = $m.RecipientTypeDetails -in $Scope.UserPrimaryTypes
            $archiveSelected = $Scope.IncludeArchives -and $m.RecipientTypeDetails -in $Scope.ArchiveOwnerTypes
        }
        $movePrimary = $primaryOnSource -and $primarySelected
        $moveArchive = $archiveOnSource -and $archiveSelected
        if ($movePrimary -and $moveArchive) { $m.MoveType = 'PrimaryAndArchive' }
        elseif ($movePrimary) { $m.MoveType = if ($m.HasArchive) { 'PrimaryOnly' } else { 'Primary' } }
        elseif ($moveArchive) { $m.MoveType = 'ArchiveOnly' }
        if ($m.MoveType) {
            $m.Status = 'ToMove'
            if ($m.MoveType -eq 'PrimaryOnly' -and $archiveOnSource) { $m.Reason = 'The archive stays on its source database (archives not selected)' }
            elseif ($m.MoveType -eq 'ArchiveOnly' -and $primaryOnSource) { $m.Reason = "The primary mailbox stays on its source database ($($m.RecipientTypeDetails) not selected)" }
        } else {
            $m.Status = 'NotSelected'
            $m.Reason = if ($primaryOnSource -and -not $primarySelected) { "$($m.RecipientTypeDetails) not selected in this execution" } else { 'Archive not selected in this execution' }
        }
    }
}

function Measure-EmmMailboxSize {
    <#
    .SYNOPSIS
        Fills the sizes of the mailboxes stored on the given databases. One Get-MailboxStatistics
        -Database call per database returns every primary and archive mailbox stored in it, which
        is much faster than one call per mailbox. Size = items + recoverable items (what is moved).
        A mailbox never opened has no statistics: its size is 0.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Mailbox, [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Database)
    $stats = @{}
    $warnings = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $Database.Count; $i++) {
        $db = $Database[$i]
        Write-Progress -Activity 'Reading mailbox sizes' -Status "$db ($($i + 1) of $($Database.Count))" -PercentComplete ([int](100 * $i / [Math]::Max(1, $Database.Count)))
        try { $rows = @(Get-MailboxStatistics -Database $db -ErrorAction Stop -WarningAction SilentlyContinue) }
        catch { $warnings.Add("Get-MailboxStatistics -Database $db`: $($_.Exception.Message)"); continue }
        foreach ($s in $rows) {
            if (Get-EmmProp $s 'DisconnectReason') { continue }
            $g = ([string](Get-EmmProp $s 'MailboxGuid' '')).ToLowerInvariant()
            if ($g) { $stats[$g] = $s }
        }
    }
    Write-Progress -Activity 'Reading mailbox sizes' -Completed
    $sizeOf = { param($s) (ConvertTo-EmmMB (Get-EmmProp $s 'TotalItemSize')) + (ConvertTo-EmmMB (Get-EmmProp $s 'TotalDeletedItemSize')) }
    foreach ($m in $Mailbox) {
        if ($m.Database -in $Database) {
            $s = if ($m.ExchangeGuid) { $stats[$m.ExchangeGuid] } else { $null }
            $m.PrimarySizeMB = if ($s) { [Math]::Round((& $sizeOf $s), 2) } else { 0.0 }
            $m.ItemCount = if ($s) { [long](Get-EmmProp $s 'ItemCount' 0) } else { 0 }
        }
        if ($m.HasArchive -and $m.ArchiveDatabase -in $Database) {
            $s = if ($m.ArchiveGuid) { $stats[$m.ArchiveGuid] } else { $null }
            $m.ArchiveSizeMB = if ($s) { [Math]::Round((& $sizeOf $s), 2) } else { 0.0 }
            $m.ArchiveItemCount = if ($s) { [long](Get-EmmProp $s 'ItemCount' 0) } else { 0 }
        }
        $size = 0.0
        if ((Test-EmmMoveType $m.MoveType 'Primary') -and $null -ne $m.PrimarySizeMB) { $size += $m.PrimarySizeMB }
        if ((Test-EmmMoveType $m.MoveType 'Archive') -and $null -ne $m.ArchiveSizeMB) { $size += $m.ArchiveSizeMB }
        $m.MoveSizeMB = [Math]::Round($size, 2)
    }
    return @($warnings)
}

function Get-EmmMailboxInventory {
    <#
    .SYNOPSIS
        Reads every mailbox of the organisation (system, user, public folder and monitoring),
        decides what this execution moves, and measures the mailboxes of the source databases.
    .OUTPUTS
        Mailboxes (normalized records), Counts per source of the list, Warnings.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$DatabaseInventory, [Parameter(Mandatory)]$Scope, [switch]$Quiet)
    $records = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    $warnings = [System.Collections.Generic.List[string]]::new()
    $lists = @(
        @{ Label = 'Arbitration'; Icon = 'System'; Params = @{ Arbitration = $true } }
        @{ Label = 'Audit log'; Icon = 'System'; Params = @{ AuditLog = $true } }
        @{ Label = 'Auxiliary audit log'; Icon = 'System'; Params = @{ AuxAuditLog = $true } }
        @{ Label = 'Discovery'; Icon = 'System'; Params = @{ RecipientTypeDetails = 'DiscoveryMailbox' } }
        @{ Label = 'User, shared, resource'; Icon = 'People'; Params = @{} }
        @{ Label = 'Public folder'; Icon = 'Folder'; Params = @{ PublicFolder = $true } }
        @{ Label = 'Monitoring (excluded)'; Icon = 'Lock'; Params = @{ Monitoring = $true }; Category = 'Monitoring' }
    )
    $counts = [ordered]@{}
    foreach ($l in $lists) {
        $p = @{}; foreach ($k in $l.Params.Keys) { $p[$k] = $l.Params[$k] }
        $p['ResultSize'] = 'Unlimited'; $p['ErrorAction'] = 'Stop'; $p['WarningAction'] = 'SilentlyContinue'
        Write-Progress -Activity 'Reading mailboxes' -Status $l.Label
        try { $items = @(Get-Mailbox @p) }
        catch {
            $warnings.Add("Get-Mailbox ($($l.Label)): $($_.Exception.Message)")
            if (-not $Quiet) { Write-EmmItem Warn ("{0}: not read ({1})" -f $l.Label, $_.Exception.Message) }
            continue
        }
        $added = 0
        foreach ($mbx in $items) {
            $r = ConvertTo-EmmMailboxRecord -Mailbox $mbx -Category $(if ($l.ContainsKey('Category')) { $l.Category } else { '' })
            $key = if ($r.Guid) { $r.Guid } elseif ($r.ExchangeGuid) { $r.ExchangeGuid } else { $r.DistinguishedName }
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true; $records.Add($r); $added++
        }
        $counts[$l.Label] = $added
        if (-not $Quiet) { Write-EmmItem Ok ('{0,-26} {1,8}' -f $l.Label, (Format-EmmNumber $added)) -Icon $l.Icon }
    }
    Write-Progress -Activity 'Reading mailboxes' -Completed
    $all = $records.ToArray()
    Set-EmmMailboxSelection -Mailbox $all -Classification $DatabaseInventory.Classification -Scope $Scope -Settings $Settings
    $toMeasure = @($all | Where-Object { $_.Category -ne 'Monitoring' -and ($_.PrimaryRole -eq 'Source' -or $_.ArchiveRole -eq 'Source') })
    foreach ($w in (Measure-EmmMailboxSize -Mailbox $toMeasure -Database $DatabaseInventory.Classification.Sources)) { $warnings.Add($w) }
    Update-EmmDatabaseCounter -Database $DatabaseInventory.Databases -Mailbox $all
    return [pscustomobject]@{ Mailboxes = $all; Counts = $counts; Warnings = @($warnings) }
}

function Update-EmmDatabaseCounter {
    <# Per database: mailboxes, archives, system and monitoring mailboxes stored in it, and what leaves it. #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Database, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Mailbox)
    $byName = @{}
    foreach ($d in $Database) { $byName[$d.Name] = $d; $d.Mailboxes = 0; $d.Archives = 0; $d.SystemMailboxes = 0; $d.MonitoringMailboxes = 0; $d.ToMove = 0; $d.ToMoveMB = 0.0 }
    foreach ($m in $Mailbox) {
        if ($m.Database -and $byName.ContainsKey($m.Database)) {
            $d = $byName[$m.Database]
            if ($m.Category -eq 'Monitoring') { $d.MonitoringMailboxes++ } elseif ($m.Category -eq 'System') { $d.SystemMailboxes++ } else { $d.Mailboxes++ }
            if ($m.Status -eq 'ToMove' -and (Test-EmmMoveType $m.MoveType 'Primary')) { $d.ToMove++; $d.ToMoveMB += [double]$m.PrimarySizeMB }
        }
        if ($m.HasArchive -and $byName.ContainsKey($m.ArchiveDatabase)) {
            $d = $byName[$m.ArchiveDatabase]; $d.Archives++
            if ($m.Status -eq 'ToMove' -and (Test-EmmMoveType $m.MoveType 'Archive')) { $d.ToMove++; $d.ToMoveMB += [double]$m.ArchiveSizeMB }
        }
    }
    foreach ($d in $Database) { $d.ToMoveMB = [Math]::Round($d.ToMoveMB, 2) }
}

function Get-EmmToolBatchName {
    <#
    .SYNOPSIS
        Name of the tool batch a move request or migration user belongs to, or '' when it does not
        belong to this tool. Move requests of a migration batch are labelled 'MigrationService:<batch>'.
    #>
    param([AllowNull()][string]$BatchName, [Parameter(Mandatory)]$Settings)
    if (-not $BatchName) { return '' }
    $name = $BatchName -replace '^MigrationService:', ''
    if ($name -eq $Settings.System.BatchName) { return $name }
    if ($name -match ('^' + [regex]::Escape($Settings.Plan.BatchNamePrefix) + '\d{2,3}$')) { return $name }
    return ''
}

function Get-EmmMigrationObjects {
    <#
    .SYNOPSIS
        Migration batches and move requests of the organisation, normalized, with the part that
        belongs to this tool (batch names Plan.BatchNamePrefix + number, and System.BatchName).
    #>
    param([Parameter(Mandatory)]$Settings)
    $batches = foreach ($b in @(Get-MigrationBatch -ErrorAction Stop -WarningAction SilentlyContinue)) {
        $name = [string](Get-EmmProp $b 'Identity' (Get-EmmProp $b 'Name' ''))
        [pscustomobject][ordered]@{
            Name = $name; Status = [string](Get-EmmProp $b 'Status' ''); IsTool = [bool](Get-EmmToolBatchName $name $Settings)
            Total = [int](Get-EmmProp $b 'TotalCount' 0); Synced = [int](Get-EmmProp $b 'SyncedCount' 0)
            Finalized = [int](Get-EmmProp $b 'FinalizedCount' 0); Failed = [int](Get-EmmProp $b 'FailedCount' 0)
        }
    }
    $moves = foreach ($r in @(Get-MoveRequest -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue)) {
        $batch = [string](Get-EmmProp $r 'BatchName' '')
        [pscustomobject][ordered]@{
            Identity = [string](Get-EmmProp $r 'Identity' ''); DisplayName = [string](Get-EmmProp $r 'DisplayName' ''); Alias = [string](Get-EmmProp $r 'Alias' '')
            ExchangeGuid = ([string](Get-EmmProp $r 'ExchangeGuid' '')).ToLowerInvariant(); Status = [string](Get-EmmProp $r 'Status' '')
            BatchName = $batch; ToolBatch = Get-EmmToolBatchName $batch $Settings; TargetDatabase = Get-EmmName (Get-EmmProp $r 'TargetDatabase' '')
        }
    }
    $batches = @($batches); $moves = @($moves)
    return [pscustomobject]@{
        Batches = $batches; MoveRequests = $moves
        ToolBatches = @($batches | Where-Object { $_.IsTool }); ToolMoves = @($moves | Where-Object { $_.ToolBatch })
        OtherBatches = @($batches | Where-Object { -not $_.IsTool }).Count; OtherMoves = @($moves | Where-Object { -not $_.ToolBatch }).Count
    }
}

function Get-EmmManagedName {
    <#
    .SYNOPSIS
        Every name managed by this tool: its migration batches, plus the labels of its move requests that
        have no migration batch (System.BatchName, and a batch made only of public folder mailboxes).
    #>
    param([Parameter(Mandatory)]$Objects)
    return @(@($Objects.ToolBatches | ForEach-Object { $_.Name }) + @($Objects.ToolMoves | ForEach-Object { $_.ToolBatch }) | Where-Object { $_ } | Sort-Object -Unique)
}

#endregion
#region 5. Plan ---------------------------------------------------------------------------------------------

function Select-EmmTargetDatabase {
    <# The least loaded database of the pool (ties: first by name); its load grows by the size placed on it. #>
    param([Parameter(Mandatory)][hashtable]$Load, [AllowEmptyCollection()][string[]]$Pool = @(), [double]$SizeMB, [string]$Purpose = 'target')
    if (-not @($Pool).Count) {
        throw $(if ($Purpose -eq 'archive') { "No mounted database matches Databases.ArchiveTargetDatabasePattern: an archive to move has no destination (archives are never placed elsewhere when this pattern is set). Fix the pattern, mount the databases, map the archive database in Databases.DatabaseMap, or leave archives out of the selection." }
            else { 'No mounted target database: nothing can be planned. Check Databases.TargetDatabasePattern and the state of the databases (-Mode Inventory).' })
    }
    $best = $null
    foreach ($name in $Pool) { if ($null -eq $best -or $Load[$name] -lt $Load[$best]) { $best = $name } }
    $Load[$best] = $Load[$best] + $SizeMB
    return $best
}

function New-EmmPlan {
    <#
    .SYNOPSIS
        Builds the migration plan from the inventory: one row per mailbox to move, with its target
        database(s) and its batch.
    .DESCRIPTION
        Targets  : DatabaseMap first; otherwise the least loaded target database (largest mailboxes
                   placed first). Archives follow DatabaseMap, then ArchiveTargetDatabasePattern, then
                   their primary mailbox (or the database of the primary when only the archive moves).
                   System mailboxes go to System.TargetDatabase when it is set.
        Batches  : user mailboxes only (system mailboxes are moved by individual move requests,
                   labelled System.BatchName). Balanced = longest processing time first with a cap
                   on the number of mailboxes per batch, so batches have equal sizes and counts.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$DatabaseInventory, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Mailbox, [Parameter(Mandatory)]$Scope, [int]$BatchCount)
    if (-not $BatchCount) { $BatchCount = $Settings.Plan.BatchCount }
    $warnings = [System.Collections.Generic.List[string]]::new()
    $dbs = @($DatabaseInventory.Databases)
    $usable = @{}
    foreach ($d in $dbs) { if (($d.IsTarget -or $d.IsArchiveTarget) -and $d.Mounted -ne $false) { $usable[$d.Name] = $d } }
    $primaryPool = @($dbs | Where-Object { $_.IsTarget -and $usable.ContainsKey($_.Name) } | ForEach-Object { $_.Name } | Sort-Object)
    $archivePool = @(if ($Settings.Databases.ArchiveTargetDatabasePattern) { $dbs | Where-Object { $_.IsArchiveTarget -and $usable.ContainsKey($_.Name) } | ForEach-Object { $_.Name } | Sort-Object } else { $primaryPool })
    $load = @{}
    foreach ($d in $dbs) { if ($usable.ContainsKey($d.Name)) { $load[$d.Name] = $(if ($Settings.Databases.CountExistingData) { [double]$d.SizeMB } else { 0.0 }) } }
    $map = $Settings.Databases.DatabaseMap
    $systemTarget = $Settings.System.TargetDatabase
    $toMove = @($Mailbox | Where-Object { $_.Status -eq 'ToMove' } | Sort-Object @{ Expression = { [double]$_.MoveSizeMB }; Descending = $true }, Name)
    if ($systemTarget -and @($toMove | Where-Object { $_.Category -eq 'System' }).Count -and $systemTarget -notin $primaryPool) {
        throw "System.TargetDatabase '$systemTarget' is not a mounted target database."
    }
    $mapped = { param($source) if ($source -and $map.Contains($source) -and $usable.ContainsKey([string]$map[$source])) { [string]$map[$source] } else { '' } }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $toMove) {
        $movesPrimary = Test-EmmMoveType $m.MoveType 'Primary'
        $movesArchive = Test-EmmMoveType $m.MoveType 'Archive'
        $primaryMB = if ($movesPrimary) { [double]$m.PrimarySizeMB } else { 0.0 }
        $archiveMB = if ($movesArchive) { [double]$m.ArchiveSizeMB } else { 0.0 }
        $target = ''; $archiveTarget = ''
        if ($movesPrimary) {
            if ($m.Category -eq 'System' -and $systemTarget) { $target = $systemTarget; $load[$target] += $primaryMB }
            else {
                $target = & $mapped $m.Database
                if ($target) { $load[$target] += $primaryMB } else { $target = Select-EmmTargetDatabase -Load $load -Pool $primaryPool -SizeMB $primaryMB }
            }
        }
        if ($movesArchive) {
            $archiveTarget = & $mapped $m.ArchiveDatabase
            if ($archiveTarget) { $load[$archiveTarget] += $archiveMB }
            elseif ($Settings.Databases.ArchiveTargetDatabasePattern) { $archiveTarget = Select-EmmTargetDatabase -Load $load -Pool $archivePool -SizeMB $archiveMB -Purpose archive }
            elseif ($movesPrimary) { $archiveTarget = $target; $load[$target] += $archiveMB }
            elseif ($usable.ContainsKey($m.Database)) { $archiveTarget = $m.Database; $load[$archiveTarget] += $archiveMB }
            else { $archiveTarget = Select-EmmTargetDatabase -Load $load -Pool $primaryPool -SizeMB $archiveMB }
        }
        $rows.Add([pscustomobject][ordered]@{
                Workload = $(if ($m.Category -eq 'System') { 'System' } else { 'User' }); Batch = ''; BatchNumber = 0
                Name = $m.Name; DisplayName = $m.DisplayName; Alias = $m.Alias; PrimarySmtpAddress = $m.PrimarySmtpAddress
                RecipientTypeDetails = $m.RecipientTypeDetails; MoveType = $m.MoveType
                SourceDatabase = $m.Database; SourceArchiveDatabase = $m.ArchiveDatabase; TargetDatabase = $target; TargetArchiveDatabase = $archiveTarget
                PrimarySizeMB = [Math]::Round($primaryMB, 2); ArchiveSizeMB = [Math]::Round($archiveMB, 2); MoveSizeMB = [Math]::Round($primaryMB + $archiveMB, 2)
                ItemCount = [long]$(if ($movesPrimary -and $null -ne $m.ItemCount) { $m.ItemCount } else { 0 }) + [long]$(if ($movesArchive -and $null -ne $m.ArchiveItemCount) { $m.ArchiveItemCount } else { 0 })
                Note = $m.Reason; Guid = $m.Guid; ExchangeGuid = $m.ExchangeGuid; ArchiveGuid = $m.ArchiveGuid; DistinguishedName = $m.DistinguishedName
            })
    }

    # ---- Batches (user mailboxes) --------------------------------------------------------------------------
    $userRows = @($rows | Where-Object { $_.Workload -eq 'User' })
    $strategy = $Settings.Plan.BatchStrategy
    $groups = @()
    if ($userRows.Count) {
        if ($strategy -eq 'Balanced') {
            # Longest processing time first: each mailbox (largest first) goes to the batch with the
            # smallest volume among those that can still take one. Exactly (n mod k) batches may hold
            # one mailbox more than the others, so the counts differ by 1 at most.
            $k = [Math]::Min($BatchCount, $userRows.Count)
            $base = [int][Math]::Floor($userRows.Count / [double]$k)
            $extra = $userRows.Count - $base * $k
            $extraUsed = 0
            $binSize = [double[]]::new($k); $binCount = [int[]]::new($k)
            foreach ($r in $userRows) {
                $best = -1
                for ($b = 0; $b -lt $k; $b++) {
                    $open = $binCount[$b] -lt $base -or ($binCount[$b] -eq $base -and $extraUsed -lt $extra)
                    if (-not $open) { continue }
                    if ($best -lt 0 -or $binSize[$b] -lt $binSize[$best] -or ($binSize[$b] -eq $binSize[$best] -and $binCount[$b] -lt $binCount[$best])) { $best = $b }
                }
                if ($binCount[$best] -eq $base) { $extraUsed++ }
                $binSize[$best] += $r.MoveSizeMB; $binCount[$best]++
                $r.BatchNumber = $best + 1
            }
        } else {
            $keyOf = if ($strategy -eq 'PerSourceDatabase') {
                { param($r) if (Test-EmmMoveType $r.MoveType 'Primary') { $r.SourceDatabase } else { $r.SourceArchiveDatabase } }
            } else {
                { param($r) if (Test-EmmMoveType $r.MoveType 'Primary') { $r.TargetDatabase } else { $r.TargetArchiveDatabase } }
            }
            $keys = @($userRows | ForEach-Object { & $keyOf $_ } | Sort-Object -Unique)
            if ($keys.Count -gt 999) { throw "The strategy $strategy gives $($keys.Count) batches (999 at most)." }
            $index = @{}; for ($i = 0; $i -lt $keys.Count; $i++) { $index[$keys[$i]] = $i + 1 }
            foreach ($r in $userRows) { $r.BatchNumber = $index[(& $keyOf $r)] }
        }
        $digits = if (@($userRows | ForEach-Object { $_.BatchNumber } | Sort-Object -Unique | Select-Object -Last 1)[0] -gt 99) { 3 } else { 2 }
        foreach ($r in $userRows) { $r.Batch = $Settings.Plan.BatchNamePrefix + ([string]$r.BatchNumber).PadLeft($digits, '0') }
    }
    foreach ($r in $rows) { if ($r.Workload -eq 'System') { $r.Batch = $Settings.System.BatchName } }

    $plan = [pscustomobject]@{ Meta = $null; Rows = $rows.ToArray(); Batches = @(); Databases = @(); Warnings = @($warnings) }
    $plan.Batches = Get-EmmPlanBatchSummary -Rows $plan.Rows -Settings $Settings
    $plan.Databases = Get-EmmPlanDatabaseSummary -Rows $plan.Rows -DatabaseInventory $DatabaseInventory
    $userBatchCount = @($plan.Batches | Where-Object { $_.Workload -eq 'User' }).Count
    $plan.Meta = [ordered]@{
        PlanId = (Get-Date).ToString('yyyyMMdd-HHmmss', $script:Invariant); Created = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)
        ToolVersion = $script:ToolVersion; Workload = $Scope.Workload; Selection = $Scope.Label; MailboxTypes = @($Scope.ExplicitTypes)
        BatchStrategy = $strategy; BatchCount = $userBatchCount; BatchNamePrefix = $Settings.Plan.BatchNamePrefix; SystemBatchName = $Settings.System.BatchName
        SourceDatabasePattern = $Settings.Databases.SourceDatabasePattern; TargetDatabasePattern = $Settings.Databases.TargetDatabasePattern
        ArchiveTargetDatabasePattern = $Settings.Databases.ArchiveTargetDatabasePattern
        Mailboxes = $plan.Rows.Count; SizeMB = [Math]::Round([double](@($plan.Rows | Measure-Object MoveSizeMB -Sum)[0].Sum), 2)
        SystemMailboxes = @($plan.Rows | Where-Object { $_.Workload -eq 'System' }).Count; UserMailboxes = $userRows.Count
        Archives = @($plan.Rows | Where-Object { Test-EmmMoveType $_.MoveType 'Archive' }).Count
        CsvDelimiter = $Settings.Report.CsvDelimiter
    }
    if ($strategy -eq 'Balanced' -and $userRows.Count -and $userRows.Count -lt $BatchCount) { $plan.Warnings += "Only $($userRows.Count) user mailbox(es): $($userRows.Count) batch(es) instead of $BatchCount." }
    return $plan
}

function Get-EmmPlanBatchSummary {
    <# One row per batch of the plan (and one for the system move requests): mailboxes, volume, types, databases. #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)]$Settings)
    $out = foreach ($g in @($Rows | Group-Object Batch | Sort-Object @{ Expression = { if ($_.Group[0].Workload -eq 'System') { 0 } else { 1 } } }, Name)) {
        $items = @($g.Group)
        $types = [ordered]@{}
        foreach ($t in @($items | Group-Object RecipientTypeDetails | Sort-Object Count -Descending)) { $types[$t.Name] = $t.Count }
        [pscustomobject][ordered]@{
            Name = $g.Name; Number = [int]$items[0].BatchNumber; Workload = $items[0].Workload; Mailboxes = $items.Count
            SizeMB = [Math]::Round([double](@($items | Measure-Object MoveSizeMB -Sum)[0].Sum), 2)
            Primaries = @($items | Where-Object { Test-EmmMoveType $_.MoveType 'Primary' }).Count
            Archives = @($items | Where-Object { Test-EmmMoveType $_.MoveType 'Archive' }).Count
            PublicFolders = @($items | Where-Object { $_.RecipientTypeDetails -eq 'PublicFolderMailbox' }).Count
            LargestMB = [Math]::Round([double](@($items | Measure-Object MoveSizeMB -Maximum)[0].Maximum), 2)
            Types = (@($types.Keys | ForEach-Object { '{0} {1}' -f $types[$_], $_ }) -join ', ')
            Sources = (@($items | ForEach-Object { if (Test-EmmMoveType $_.MoveType 'Primary') { $_.SourceDatabase }; if (Test-EmmMoveType $_.MoveType 'Archive') { $_.SourceArchiveDatabase } } | Sort-Object -Unique) -join ', ')
            Targets = (@($items | ForEach-Object { $_.TargetDatabase; $_.TargetArchiveDatabase } | Where-Object { $_ } | Sort-Object -Unique) -join ', ')
        }
    }
    return @($out)
}

function Get-EmmPlanDatabaseSummary {
    <# Target and source databases of the plan: existing volume, planned arrivals and departures. #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)]$DatabaseInventory)
    $byName = @{}
    foreach ($d in $DatabaseInventory.Databases) { $byName[$d.Name] = $d; $d.PlannedPrimaries = 0; $d.PlannedArchives = 0; $d.PlannedMB = 0.0 }
    foreach ($r in $Rows) {
        if ($r.TargetDatabase -and $byName.ContainsKey($r.TargetDatabase)) { $d = $byName[$r.TargetDatabase]; $d.PlannedPrimaries++; $d.PlannedMB += $r.PrimarySizeMB }
        if ($r.TargetArchiveDatabase -and $byName.ContainsKey($r.TargetArchiveDatabase)) { $d = $byName[$r.TargetArchiveDatabase]; $d.PlannedArchives++; $d.PlannedMB += $r.ArchiveSizeMB }
    }
    $out = foreach ($d in $DatabaseInventory.Databases) {
        if (-not ($d.IsSource -or $d.IsTarget -or $d.IsArchiveTarget)) { continue }
        [pscustomobject][ordered]@{
            Name = $d.Name; Role = $d.Role; Server = $d.Server; Mounted = $d.Mounted; ExistingMB = $d.SizeMB
            PlannedPrimaries = $d.PlannedPrimaries; PlannedArchives = $d.PlannedArchives; PlannedMB = [Math]::Round($d.PlannedMB, 2)
            LeavingMailboxes = $d.ToMove; LeavingMB = $d.ToMoveMB; MapTarget = $d.MapTarget
        }
    }
    return @($out)
}

$script:PlanColumns = @('Workload', 'Batch', 'BatchNumber', 'Name', 'DisplayName', 'Alias', 'PrimarySmtpAddress', 'RecipientTypeDetails', 'MoveType',
    'SourceDatabase', 'SourceArchiveDatabase', 'TargetDatabase', 'TargetArchiveDatabase', 'PrimarySizeMB', 'ArchiveSizeMB', 'MoveSizeMB', 'ItemCount',
    'Note', 'Guid', 'ExchangeGuid', 'ArchiveGuid', 'DistinguishedName')

function Save-EmmPlan {
    <#
    .SYNOPSIS
        Writes the plan in a folder: MigrationPlan.csv (the rows, read back by -Mode Start) and
        MigrationPlan.json (description, batches and databases).
    #>
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)]$Settings)
    [void][IO.Directory]::CreateDirectory($Directory)
    $csv = Join-Path $Directory 'MigrationPlan.csv'
    Export-EmmCsv -Rows $Plan.Rows -Columns $script:PlanColumns -Path $csv -Delimiter $Settings.Report.CsvDelimiter
    $json = Join-Path $Directory 'MigrationPlan.json'
    $doc = [ordered]@{ Meta = $Plan.Meta; Batches = @($Plan.Batches); Databases = @($Plan.Databases); Warnings = @($Plan.Warnings) }
    [IO.File]::WriteAllText($json, (ConvertTo-EmmJson $doc), (New-Object Text.UTF8Encoding($false)))
    return @($csv, $json)
}

function Import-EmmPlan {
    <# Reads a plan written by Save-EmmPlan (folder, .json or .csv path). #>
    param([Parameter(Mandatory)][string]$Path)
    $dir = if (Test-Path -LiteralPath $Path -PathType Container) { $Path } else { Split-Path $Path -Parent }
    $json = Join-Path $dir 'MigrationPlan.json'; $csv = Join-Path $dir 'MigrationPlan.csv'
    if (-not (Test-Path -LiteralPath $json) -or -not (Test-Path -LiteralPath $csv)) { throw "No plan in $dir (MigrationPlan.json and MigrationPlan.csv are expected)." }
    $doc = [IO.File]::ReadAllText($json) | ConvertFrom-Json
    $delimiter = [string](Get-EmmProp $doc.Meta 'CsvDelimiter' ';')
    $rows = @(Import-Csv -LiteralPath $csv -Delimiter $delimiter -Encoding UTF8)
    foreach ($r in $rows) {
        foreach ($p in $r.PSObject.Properties) { if ($p.Value -is [string] -and $p.Value -match "^'[=+\-@]") { $p.Value = $p.Value.Substring(1) } }
        $r.BatchNumber = [int]$r.BatchNumber
        foreach ($n in 'PrimarySizeMB', 'ArchiveSizeMB', 'MoveSizeMB') { $r.$n = [double]::Parse(([string]$r.$n -replace ',', '.'), $script:Invariant) }
        $r.ItemCount = [long]$r.ItemCount
    }
    foreach ($c in $script:PlanColumns) { if ($rows.Count -and -not $rows[0].PSObject.Properties[$c]) { throw "The plan $csv has no column '$c': it was not written by this version of the tool." } }
    $meta = [ordered]@{}
    foreach ($p in $doc.Meta.PSObject.Properties) { $meta[$p.Name] = $p.Value }
    return [pscustomobject]@{ Meta = $meta; Rows = $rows; Batches = @($doc.Batches); Databases = @($doc.Databases); Warnings = @($doc.Warnings); Directory = (Resolve-Path -LiteralPath $dir).Path }
}

function Find-EmmPlan {
    <#
    .SYNOPSIS
        Most recent plan of the output folder that covers the workload (a plan 'All' covers System and User).
    #>
    param([Parameter(Mandatory)][string]$OutputPath, [Parameter(Mandatory)][string]$Workload)
    if (-not (Test-Path -LiteralPath $OutputPath)) { return $null }
    $dirs = @(Get-ChildItem -LiteralPath $OutputPath -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{6}_Plan' } | Sort-Object Name -Descending)
    foreach ($d in $dirs) {
        $json = Join-Path $d.FullName 'MigrationPlan.json'
        if (-not (Test-Path -LiteralPath $json) -or -not (Test-Path -LiteralPath (Join-Path $d.FullName 'MigrationPlan.csv'))) { continue }
        try { $meta = ([IO.File]::ReadAllText($json) | ConvertFrom-Json).Meta } catch { continue }
        $w = [string]$meta.Workload
        if ($w -eq 'All' -or $w -eq $Workload) { return $d.FullName }
    }
    return $null
}

#endregion
#region 6. Migration actions ---------------------------------------------------------------------------------

function Reset-EmmActions { $script:Actions.Clear() }
function Get-EmmActions { return @($script:Actions.ToArray()) }

function Add-EmmAction {
    <#
    .SYNOPSIS
        Records the result of one action (report + log) and shows it in the console unless -Quiet.
        Status: Success | Simulated | Skipped | Failed | AlreadyDone.
    #>
    param([string]$Workload = '', [string]$Batch = '', [Parameter(Mandatory)][string]$Target, [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][ValidateSet('Success', 'Simulated', 'Skipped', 'Failed', 'AlreadyDone')][string]$Status, [string]$Detail = '', [switch]$Quiet)
    $script:Actions.Add([pscustomobject][ordered]@{
            Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant); Workload = $Workload; Batch = $Batch
            Target = $Target; Action = $Action; Status = $Status; Detail = $Detail
        })
    $text = '{0}  {1}{2}' -f $Target, $Action, $(if ($Detail) { "  $($script:Dot) $Detail" } else { '' })
    if ($Quiet) {
        Write-EmmLog $(if ($Status -eq 'Failed') { 'ERROR' } elseif ($Status -eq 'Skipped') { 'WARN' } else { 'OK' }) ("[$Status] $text")
    } else {
        $item = @{ Success = 'Ok'; Simulated = 'Info'; Skipped = 'Skip'; Failed = 'Fail'; AlreadyDone = 'Ok' }[$Status]
        Write-EmmItem $item $text -Icon $(if ($Status -eq 'Simulated') { 'Simulate' } else { '' })
    }
}

function Invoke-EmmChange {
    <#
    .SYNOPSIS
        Runs one Exchange command that changes something. The command line is written to the log
        (audit trail); in simulation the command runs with -WhatIf: Exchange checks it, nothing changes.
    #>
    param([Parameter(Mandatory)][string]$Command, [hashtable]$Parameters = @{}, [switch]$Simulate)
    $p = @{}
    foreach ($k in $Parameters.Keys) { $p[$k] = $Parameters[$k] }
    $p['ErrorAction'] = 'Stop'; $p['WarningAction'] = 'SilentlyContinue'; $p['Confirm'] = $false
    if ($Simulate) { $p['WhatIf'] = $true }
    $shown = foreach ($k in @($Parameters.Keys | Sort-Object)) {
        $v = $Parameters[$k]
        if ($v -is [bool] -or $v -is [switch]) { if ($v) { "-$k" } else { "-$k`:`$false" } }
        elseif ($v -is [byte[]]) { "-$k <$($v.Length) bytes>" }
        elseif ($v -is [datetime]) { "-$k '$($v.ToString('yyyy-MM-dd HH:mm', $script:Invariant))'" }
        elseif ($v -is [array]) { "-$k " + (@($v | ForEach-Object { "'$_'" }) -join ',') }
        else { "-$k '$v'" }
    }
    Write-EmmLog 'CHANGE' ('{0}{1} {2}' -f $(if ($Simulate) { '[WhatIf] ' } else { '' }), $Command, ($shown -join ' '))
    return (& $Command @p)
}

function Get-EmmCurrentMailbox {
    <# Current state of the mailboxes of the plan (bulk reads), indexed by ExchangeGuid and Guid. #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)
    $index = @{}
    $lists = @()
    if (@($Rows | Where-Object { $_.Workload -eq 'System' }).Count) {
        $lists += @(@{ Arbitration = $true }, @{ AuditLog = $true }, @{ AuxAuditLog = $true }, @{ RecipientTypeDetails = 'DiscoveryMailbox' })
    }
    if (@($Rows | Where-Object { $_.Workload -eq 'User' -and $_.RecipientTypeDetails -ne 'PublicFolderMailbox' }).Count) { $lists += @(@{}) }
    if (@($Rows | Where-Object { $_.RecipientTypeDetails -eq 'PublicFolderMailbox' }).Count) { $lists += @(@{ PublicFolder = $true }) }
    foreach ($l in $lists) {
        $p = @{}; foreach ($k in $l.Keys) { $p[$k] = $l[$k] }
        $p['ResultSize'] = 'Unlimited'; $p['ErrorAction'] = 'Stop'; $p['WarningAction'] = 'SilentlyContinue'
        try { $items = @(Get-Mailbox @p) } catch { continue }
        foreach ($mbx in $items) {
            $r = ConvertTo-EmmMailboxRecord -Mailbox $mbx
            if ($r.ExchangeGuid) { $index[$r.ExchangeGuid] = $r }
            if ($r.Guid) { $index[$r.Guid] = $r }
        }
    }
    return $index
}

function Get-EmmStartPreflight {
    <#
    .SYNOPSIS
        Checks every row of the plan against the current state of the organisation and decides
        what -Mode Start will do. Nothing is changed here.
    .DESCRIPTION
        For each mailbox: Submit, or Skip with the reason (status Skipped, AlreadyDone or Failed).
          - mailbox deleted, or moved elsewhere since the plan          -> Skipped
          - already on its target database(s)                           -> AlreadyDone
          - part already moved (primary or archive)                     -> the move type is reduced
          - target database missing, dismounted or no longer a target   -> Failed
          - finished move request (completed / failed)                  -> removed first (Move.ReplaceFinishedMoveRequests)
          - move request in progress                                    -> Skipped
          - migration user of an earlier batch of the tool              -> that batch removed first when all its moves are
                                                                           finished (Test-EmmBatchRemovable, the rule of
                                                                           Cleanup), otherwise Skipped; orphan user removed
          - migration user of a batch of another origin                 -> Skipped
          - monitoring mailbox                                          -> Failed (never moved)
        For each user batch: Create, Replace (a batch of the tool with the same name whose moves are all finished)
        or Skip (not finished, or not a batch of the tool).
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][ValidateSet('System', 'User', 'All')][string]$Workload)
    # The names of the plan must be names of this tool with the CURRENT configuration, otherwise its batches
    # and move requests would not be recognised (and protected) as the tool's own afterwards.
    foreach ($pair in @(@('BatchNamePrefix', $Settings.Plan.BatchNamePrefix, 'Plan.BatchNamePrefix'), @('SystemBatchName', $Settings.System.BatchName, 'System.BatchName'))) {
        $planned = [string]$Plan.Meta[$pair[0]]
        if ($planned -and $planned -ne $pair[1]) { throw "The plan was made with $($pair[2]) = '$planned'; the configuration now says '$($pair[1])'. Create a new plan (or restore the setting)." }
    }
    $rows = @($Plan.Rows | Where-Object { $Workload -eq 'All' -or $_.Workload -eq $Workload })
    $foreign = @($rows | Where-Object { ($_.Workload -eq 'System' -and $_.Batch -ne $Settings.System.BatchName) -or ($_.Workload -ne 'System' -and -not (Get-EmmToolBatchName $_.Batch $Settings)) -or ($_.Workload -ne 'System' -and $_.Batch -eq $Settings.System.BatchName) })
    if ($foreign.Count) { throw ("The plan has {0} row(s) with a batch name that is not a name of this tool (e.g. '{1}'): the plan file was edited or made with another configuration. Create a new plan." -f $foreign.Count, $foreign[0].Batch) }
    $warnings = [System.Collections.Generic.List[string]]::new()

    $dbState = @{}
    foreach ($db in @(Get-MailboxDatabase -Status -ErrorAction Stop -WarningAction SilentlyContinue)) { $dbState[(Get-EmmName $db.Name)] = Get-EmmProp $db 'Mounted' }
    $class = Get-EmmDatabaseClassification -Name @($dbState.Keys) -Settings $Settings
    $current = Get-EmmCurrentMailbox -Rows $rows
    $objects = Get-EmmMigrationObjects -Settings $Settings
    $movesByGuid = @{}
    foreach ($mr in $objects.MoveRequests) { if ($mr.ExchangeGuid) { $movesByGuid[$mr.ExchangeGuid] = $mr } }
    $batchesByName = @{}
    foreach ($b in $objects.Batches) { $batchesByName[$b.Name] = $b }
    $usersByKey = @{}
    if (@($rows | Where-Object { $_.Workload -eq 'User' }).Count) {
        foreach ($u in @(Get-MigrationUser -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue)) {
            $info = [pscustomobject]@{ Identity = [string](Get-EmmProp $u 'Identity' ''); Batch = [string](Get-EmmProp $u 'BatchId' ''); Status = [string](Get-EmmProp $u 'Status' '') }
            $g = ([string](Get-EmmProp $u 'MailboxGuid' '')).ToLowerInvariant()
            if ($g) { $usersByKey[$g] = $info }
            $mail = ([string](Get-EmmProp $u 'MailboxEmailAddress' (Get-EmmProp $u 'Identity' ''))).ToLowerInvariant()
            if ($mail) { $usersByKey[$mail] = $info }
        }
    }
    # Mailboxes about to be moved again: their failed moves in an earlier batch do not block its removal.
    $retried = @{}
    foreach ($r in $rows) { if ($r.ExchangeGuid) { $retried[$r.ExchangeGuid] = $true } }
    $removable = @{}
    $batchRemovable = {
        param([string]$Name)
        if (-not $removable.ContainsKey($Name)) {
            $removable[$Name] = Test-EmmBatchRemovable -Batch $batchesByName[$Name] -Moves @($objects.ToolMoves | Where-Object { $_.ToolBatch -eq $Name }) -Settings $Settings -RetriedGuid $retried
        }
        return $removable[$Name]
    }

    # ---- Batches ----------------------------------------------------------------------------------------
    $batchDecision = @{}
    foreach ($name in @($rows | Where-Object { $_.Workload -eq 'User' } | ForEach-Object { $_.Batch } | Sort-Object -Unique)) {
        $existing = if ($batchesByName.ContainsKey($name)) { $batchesByName[$name] } else { $null }
        $test = if ($existing -and $existing.IsTool) { & $batchRemovable $name } else { $null }
        $batchDecision[$name] = if (-not $existing) { [pscustomobject]@{ Name = $name; Action = 'Create'; Existing = ''; Reason = '' } }
            elseif (-not $existing.IsTool) { [pscustomobject]@{ Name = $name; Action = 'Skip'; Existing = $existing.Status; Reason = "A batch named $name exists and does not belong to this tool: it is never removed. Use another Plan.BatchNamePrefix and create a new plan" } }
            elseif ($test.Removable) { [pscustomobject]@{ Name = $name; Action = 'Replace'; Existing = $existing.Status; Reason = "Earlier batch with the same name removed first ($($test.Reason))" } }
            else { [pscustomobject]@{ Name = $name; Action = 'Skip'; Existing = $existing.Status; Reason = "Batch $name already exists and is not finished ($($test.Reason)): follow it with -Mode Status, or use another Plan.BatchNamePrefix" } }
    }
    # Earlier batches of the tool (other names) that still hold planned mailboxes and are removed first.
    $previous = @{}

    # ---- Mailboxes ---------------------------------------------------------------------------------------
    # A primary mailbox goes to a target database; an archive to a target or an archive target database
    # (only an archive target database when Databases.ArchiveTargetDatabasePattern is set).
    $primaryOk = { param($db) $db -and $class.Roles.ContainsKey($db) -and $class.Roles[$db].IsTarget }
    $archiveOk = { param($db) $db -and $class.Roles.ContainsKey($db) -and ($class.Roles[$db].IsArchiveTarget -or (-not $Settings.Databases.ArchiveTargetDatabasePattern -and $class.Roles[$db].IsTarget) -or ($Settings.Databases.DatabaseMap.Values -contains $db)) }
    $items = foreach ($r in $rows) {
        $decision = 'Submit'; $status = ''; $reason = ''; $removeMove = ''; $removeUser = ''; $previousBatch = ''
        $moveType = $r.MoveType
        $cur = $null
        if ($r.ExchangeGuid -and $current.ContainsKey($r.ExchangeGuid)) { $cur = $current[$r.ExchangeGuid] } elseif ($r.Guid -and $current.ContainsKey($r.Guid)) { $cur = $current[$r.Guid] }
        $needPrimary = Test-EmmMoveType $moveType 'Primary'
        $needArchive = Test-EmmMoveType $moveType 'Archive'
        if (-not $cur) { $decision = 'Skip'; $status = 'Skipped'; $reason = 'Mailbox not found (removed since the plan?)' }
        elseif ($cur.Category -eq 'Monitoring') { $decision = 'Skip'; $status = 'Failed'; $reason = 'Monitoring mailbox: never moved' }
        else {
            $notes = @()
            if ($needPrimary -and $cur.Database -ne $r.SourceDatabase) {
                if (& $primaryOk $cur.Database) { $needPrimary = $false; $notes += "primary already on $($cur.Database)" }
                else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Primary mailbox moved since the plan (now on $($cur.Database)): create a new plan" }
            }
            if ($decision -eq 'Submit' -and $needArchive -and $cur.ArchiveDatabase -ne $r.SourceArchiveDatabase) {
                if (& $archiveOk $cur.ArchiveDatabase) { $needArchive = $false; $notes += "archive already on $($cur.ArchiveDatabase)" }
                else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Archive moved since the plan (now on $($cur.ArchiveDatabase)): create a new plan" }
            }
            if ($decision -eq 'Submit' -and -not $needPrimary -and -not $needArchive) { $decision = 'Skip'; $status = 'AlreadyDone'; $reason = 'Already on the target database(s)' }
            if ($decision -eq 'Submit') {
                $moveType = if ($needPrimary -and $needArchive) { 'PrimaryAndArchive' } elseif ($needPrimary) { if ($cur.HasArchive) { 'PrimaryOnly' } else { 'Primary' } } else { 'ArchiveOnly' }
                if ($notes.Count) { $reason = 'Move type reduced to ' + $moveType + ' (' + ($notes -join ', ') + ')' }
                foreach ($check in @(@($needPrimary, $r.TargetDatabase, $primaryOk, 'target'), @($needArchive, $r.TargetArchiveDatabase, $archiveOk, 'archive target'))) {
                    if (-not $check[0] -or $decision -ne 'Submit') { continue }
                    $t = $check[1]
                    if (-not $t -or -not $dbState.ContainsKey($t) -or $dbState[$t] -eq $false -or -not (& $check[2] $t)) {
                        $decision = 'Skip'; $status = 'Failed'; $reason = "The $($check[3]) database '$t' is not available (missing, dismounted or no longer a $($check[3]) in the configuration): create a new plan"
                    }
                }
            }
            if ($decision -eq 'Submit' -and $r.ExchangeGuid -and $movesByGuid.ContainsKey($r.ExchangeGuid)) {
                $mr = $movesByGuid[$r.ExchangeGuid]
                $owner = if ($mr.ToolBatch) { "this tool, $($mr.ToolBatch)" } elseif ($mr.BatchName) { $mr.BatchName } else { 'no batch name: another origin' }
                if ($mr.Status -in $script:FinishedMoveStatuses) {
                    $policy = $Settings.Move.ReplaceFinishedMoveRequests
                    if ($policy -eq 'All' -or ($policy -eq 'Tool' -and $mr.ToolBatch)) { $removeMove = $mr.Identity }
                    else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Finished move request ($($mr.Status), $owner) not removed by the tool (Move.ReplaceFinishedMoveRequests = $policy): remove it (Remove-MoveRequest) or set 'All'" }
                } else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Move request in progress ($($mr.Status), $owner)" }
            }
            if ($decision -eq 'Submit' -and $r.Workload -eq 'User') {
                $bd = $batchDecision[$r.Batch]
                $mu = $null
                if ($r.ExchangeGuid -and $usersByKey.ContainsKey($r.ExchangeGuid)) { $mu = $usersByKey[$r.ExchangeGuid] }
                elseif ($r.PrimarySmtpAddress -and $usersByKey.ContainsKey($r.PrimarySmtpAddress.ToLowerInvariant())) { $mu = $usersByKey[$r.PrimarySmtpAddress.ToLowerInvariant()] }
                if ($bd.Action -eq 'Skip') { $decision = 'Skip'; $status = 'Skipped'; $reason = $bd.Reason }
                elseif ($mu -and -not ($mu.Batch -eq $r.Batch -and $bd.Action -eq 'Replace')) {
                    # Exchange refuses a mailbox that is still a migration user of another batch: free it when it is safe.
                    $old = $mu.Batch
                    if (-not $old -or -not $batchesByName.ContainsKey($old)) {
                        if ($old -and (Get-EmmToolBatchName $old $Settings)) { $removeUser = $mu.Identity; $notes += "orphan migration user of $old removed first" }
                        else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Migration user left by '$(if ($old) { $old } else { 'no batch' })', which is not a batch of this tool: remove it (Remove-MigrationUser), then start again" }
                    } elseif (-not $batchesByName[$old].IsTool) {
                        $decision = 'Skip'; $status = 'Skipped'; $reason = "Migration user of the batch $old, which does not belong to this tool: it is never removed"
                    } else {
                        $t = & $batchRemovable $old
                        if ($t.Removable) {
                            $previousBatch = $old
                            if (-not $previous.ContainsKey($old)) { $previous[$old] = [pscustomobject]@{ Name = $old; Status = $batchesByName[$old].Status; Reason = $t.Reason; Mailboxes = 0 } }
                            $previous[$old].Mailboxes++
                            $notes += "earlier batch $old removed first ($($t.Reason))"
                        } else { $decision = 'Skip'; $status = 'Skipped'; $reason = "Still a migration user of the batch $old ($($t.Reason)): let it finish or complete it, then start again" }
                    }
                }
                if ($decision -eq 'Submit' -and $notes.Count) { $reason = $(if ($reason) { $reason + '; ' } else { '' }) + (@($notes | Where-Object { $_ -notlike '*already on*' }) -join '; ') }
            }
        }
        [pscustomobject]@{ Row = $r; Decision = $decision; Status = $status; Reason = $reason.Trim('; '); MoveType = $moveType; RemoveMoveRequest = $removeMove; RemoveMigrationUser = $removeUser; PreviousBatch = $previousBatch }
    }
    $items = @($items)
    if ($Workload -eq 'User') {
        $arbitration = @()
        try { $arbitration = @(Get-Mailbox -Arbitration -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue) } catch { }
        $left = @($arbitration | Where-Object { $db = Get-EmmName (Get-EmmProp $_ 'Database' ''); $class.Roles.ContainsKey($db) -and $class.Roles[$db].IsSource }).Count
        if ($left) { $warnings.Add("$left arbitration mailbox(es) are still on source databases. Recommended order: move the system mailboxes first (-Workload System).") }
    }
    if ($Workload -in 'System', 'All' -and @($rows | Where-Object { $_.Workload -eq 'System' }).Count) {
        $active = @($objects.ToolBatches | Where-Object { $_.Status -notin $script:FinishedBatchStatuses })
        if ($active.Count) { $warnings.Add("$($active.Count) migration batch(es) of this tool are active: the arbitration mailboxes (including the one used by the migration service) are best moved before or after the user batches, not during.") }
    }
    if ($Settings.Move.LargeItemLimit -gt 0 -and @($rows | Where-Object { $_.Workload -eq 'User' -and $_.RecipientTypeDetails -ne 'PublicFolderMailbox' }).Count) {
        $warnings.Add("Move.LargeItemLimit ($($Settings.Move.LargeItemLimit)) applies to the individual move requests only: Exchange has no large item limit for local migration batches (a user mailbox with a large item fails and is reported).")
    }
    return [pscustomobject]@{
        Items = $items; Batches = @($batchDecision.Values | Sort-Object Name); PreviousBatches = @($previous.Values | Sort-Object Name); Warnings = @($warnings); Workload = $Workload
        Submit = @($items | Where-Object { $_.Decision -eq 'Submit' }).Count
        Skip = @($items | Where-Object { $_.Decision -eq 'Skip' }).Count
    }
}

function Get-EmmMoveParameter {
    <# Common New-MoveRequest parameters (limits, domain controller), shared by system and public folder moves. #>
    param([Parameter(Mandatory)]$Settings, [switch]$AllowLargeItems)
    $p = @{ BadItemLimit = $Settings.Move.BadItemLimit }
    if ($AllowLargeItems) { $p['AllowLargeItems'] = $true } elseif ($Settings.Move.LargeItemLimit -gt 0) { $p['LargeItemLimit'] = $Settings.Move.LargeItemLimit }
    if ($Settings.Move.AcceptLargeDataLoss) { $p['AcceptLargeDataLoss'] = $true }
    if ($Settings.Connection.DomainController) { $p['DomainController'] = $Settings.Connection.DomainController }
    return $p
}

function Start-EmmMigration {
    <#
    .SYNOPSIS
        Executes the decisions of Get-EmmStartPreflight.
    .DESCRIPTION
        1. Finished move requests of the planned mailboxes are removed (when allowed).
        2. System mailboxes: one New-MoveRequest each, labelled System.BatchName, completed
           automatically by Exchange; optionally waited for (System.WaitForCompletion).
        3. User mailboxes: one local migration batch per plan batch (New-MigrationBatch -Local,
           CSV with TargetDatabase, TargetArchiveDatabase and MailboxType), then Start-MigrationBatch.
           The batch stops when synchronised: completion is decided later (-Mode Complete).
           Public folder mailboxes cannot be in a local batch: they get a New-MoveRequest labelled
           with the batch name and suspended when ready to complete, so -Mode Complete finishes them too.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$Preflight, [switch]$Simulate)
    $okStatus = if ($Simulate) { 'Simulated' } else { 'Success' }
    foreach ($i in @($Preflight.Items | Where-Object { $_.Decision -eq 'Skip' })) {
        Add-EmmAction -Workload $i.Row.Workload -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Not submitted' -Status $i.Status -Detail $i.Reason -Quiet
    }
    $submit = @($Preflight.Items | Where-Object { $_.Decision -eq 'Submit' })

    # ---- 1. Finished move requests ------------------------------------------------------------------------
    foreach ($i in @($submit | Where-Object { $_.RemoveMoveRequest })) {
        try {
            [void](Invoke-EmmChange -Command 'Remove-MoveRequest' -Parameters @{ Identity = $i.RemoveMoveRequest } -Simulate:$Simulate)
            Add-EmmAction -Workload $i.Row.Workload -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Remove-MoveRequest (finished request)' -Status $okStatus -Quiet
        } catch {
            $i.Decision = 'Skip'
            Add-EmmAction -Workload $i.Row.Workload -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Remove-MoveRequest (finished request)' -Status Failed -Detail $_.Exception.Message
        }
    }

    # ---- 1b. Earlier batches of the tool that still hold planned mailboxes, and orphan migration users ----------
    # Exchange refuses a mailbox that is still a migration user of another batch. These batches were checked by
    # the pre-flight with the same rule as -Mode Cleanup (all their moves finished: nothing is lost).
    $removedBatches = @{}
    foreach ($pb in @($Preflight.PSObject.Properties['PreviousBatches'] | ForEach-Object { $_.Value } | Where-Object { $_ })) {
        $dependents = @($submit | Where-Object { $_.PreviousBatch -eq $pb.Name -and $_.Decision -eq 'Submit' })
        if (-not $dependents.Count) { continue }
        try {
            [void](Invoke-EmmChange -Command 'Remove-MigrationBatch' -Parameters @{ Identity = $pb.Name; Force = $true } -Simulate:$Simulate)
            $removedBatches[$pb.Name] = $true
            Add-EmmAction -Workload 'User' -Batch $pb.Name -Target $pb.Name -Action 'Remove-MigrationBatch (earlier batch of planned mailboxes)' -Status $okStatus -Detail ('{0}; {1} planned mailbox(es) in it' -f $pb.Reason, $dependents.Count)
        } catch {
            Add-EmmAction -Workload 'User' -Batch $pb.Name -Target $pb.Name -Action 'Remove-MigrationBatch (earlier batch of planned mailboxes)' -Status Failed -Detail $_.Exception.Message
            foreach ($i in $dependents) { $i.Decision = 'Skip'; Add-EmmAction -Workload 'User' -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Not submitted' -Status Failed -Detail "The earlier batch $($pb.Name) could not be removed" -Quiet }
        }
    }
    # The migration service forgets the users of a removed batch with a delay.
    if ($removedBatches.Count -and -not $Simulate -and $Settings.Cleanup.SettleSeconds) { Start-Sleep -Seconds $Settings.Cleanup.SettleSeconds }
    foreach ($i in @($submit | Where-Object { $_.RemoveMigrationUser -and $_.Decision -eq 'Submit' })) {
        try {
            [void](Invoke-EmmChange -Command 'Remove-MigrationUser' -Parameters @{ Identity = $i.RemoveMigrationUser } -Simulate:$Simulate)
            Add-EmmAction -Workload 'User' -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Remove-MigrationUser (orphan of an earlier batch)' -Status $okStatus -Quiet
        } catch {
            $i.Decision = 'Skip'
            Add-EmmAction -Workload 'User' -Batch $i.Row.Batch -Target $i.Row.PrimarySmtpAddress -Action 'Remove-MigrationUser (orphan of an earlier batch)' -Status Failed -Detail $_.Exception.Message
        }
    }

    # ---- 2. System mailboxes -------------------------------------------------------------------------------
    $systemItems = @($submit | Where-Object { $_.Row.Workload -eq 'System' -and $_.Decision -eq 'Submit' })
    $submittedSystem = 0
    foreach ($i in $systemItems) {
        $r = $i.Row
        $p = Get-EmmMoveParameter -Settings $Settings -AllowLargeItems:$Settings.System.AllowLargeItems
        $p['Identity'] = $(if ($r.DistinguishedName) { $r.DistinguishedName } else { $r.Guid })
        $p['BatchName'] = $Settings.System.BatchName
        if (Test-EmmMoveType $i.MoveType 'Primary') { $p['TargetDatabase'] = $r.TargetDatabase }
        if (Test-EmmMoveType $i.MoveType 'Archive') { $p['ArchiveTargetDatabase'] = $r.TargetArchiveDatabase }
        if ($i.MoveType -eq 'PrimaryOnly') { $p['PrimaryOnly'] = $true } elseif ($i.MoveType -eq 'ArchiveOnly') { $p['ArchiveOnly'] = $true }
        $target = if ($r.TargetDatabase) { $r.TargetDatabase } else { $r.TargetArchiveDatabase }
        try {
            [void](Invoke-EmmChange -Command 'New-MoveRequest' -Parameters $p -Simulate:$Simulate)
            $submittedSystem++
            Add-EmmAction -Workload 'System' -Batch $r.Batch -Target $r.Name -Action "New-MoveRequest $($script:Arrow) $target" -Status $okStatus -Detail $r.RecipientTypeDetails
        } catch {
            Add-EmmAction -Workload 'System' -Batch $r.Batch -Target $r.Name -Action "New-MoveRequest $($script:Arrow) $target" -Status Failed -Detail $_.Exception.Message
        }
    }
    if ($submittedSystem -and -not $Simulate -and $Settings.System.WaitForCompletion) { Wait-EmmSystemMove -Settings $Settings }

    # ---- 3. User batches ------------------------------------------------------------------------------------
    foreach ($bd in $Preflight.Batches) {
        $name = $bd.Name
        $items = @($submit | Where-Object { $_.Row.Workload -eq 'User' -and $_.Row.Batch -eq $name -and $_.Decision -eq 'Submit' })
        if ($bd.Action -eq 'Skip') { Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'New-MigrationBatch' -Status Skipped -Detail $bd.Reason; continue }
        if (-not $items.Count) { Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'New-MigrationBatch' -Status Skipped -Detail 'No mailbox left to submit in this batch'; continue }
        if ($bd.Action -eq 'Replace' -and -not $removedBatches.ContainsKey($name)) {
            try {
                [void](Invoke-EmmChange -Command 'Remove-MigrationBatch' -Parameters @{ Identity = $name; Force = $true } -Simulate:$Simulate)
                Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'Remove-MigrationBatch (finished batch)' -Status $okStatus -Detail "previous status $($bd.Existing)"
            } catch {
                Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'Remove-MigrationBatch (finished batch)' -Status Failed -Detail $_.Exception.Message
                foreach ($i in $items) { Add-EmmAction -Workload 'User' -Batch $name -Target $i.Row.PrimarySmtpAddress -Action 'Not submitted' -Status Failed -Detail 'The previous batch with the same name could not be removed' -Quiet }
                continue
            }
        }
        $mailboxItems = @($items | Where-Object { $_.Row.RecipientTypeDetails -ne 'PublicFolderMailbox' })
        $folderItems = @($items | Where-Object { $_.Row.RecipientTypeDetails -eq 'PublicFolderMailbox' })
        $volume = Format-EmmSize ([double](@($items | ForEach-Object { $_.Row.MoveSizeMB } | Measure-Object -Sum)[0].Sum))
        if ($mailboxItems.Count) {
            $lines = [System.Collections.Generic.List[string]]::new()
            $lines.Add('EmailAddress,TargetDatabase,TargetArchiveDatabase,MailboxType')
            foreach ($i in $mailboxItems) {
                $type = if ($i.MoveType -eq 'Primary') { '' } else { $i.MoveType }
                $tp = if (Test-EmmMoveType $i.MoveType 'Primary') { $i.Row.TargetDatabase } else { '' }
                $ta = if (Test-EmmMoveType $i.MoveType 'Archive') { $i.Row.TargetArchiveDatabase } else { '' }
                $lines.Add(('{0},{1},{2},{3}' -f $i.Row.PrimarySmtpAddress, $tp, $ta, $type))
            }
            # The Local parameter set of New-MigrationBatch has no LargeItemLimit (only individual move requests have it).
            $p = @{ Name = $name; Local = $true; CSVData = [Text.Encoding]::UTF8.GetBytes(($lines -join "`r`n")); BadItemLimit = $Settings.Move.BadItemLimit }
            if ($Settings.Move.NotificationEmails.Count) { $p['NotificationEmails'] = @($Settings.Move.NotificationEmails) }
            $batchDetail = "{0} mailbox(es), {1}" -f $mailboxItems.Count, $volume
            # In simulation the batches and users to remove first still exist: Exchange would refuse the creation.
            $dependsOnRemoval = $bd.Action -eq 'Replace' -or @($mailboxItems | Where-Object { $_.PreviousBatch -or $_.RemoveMigrationUser }).Count -gt 0
            try {
                if ($Simulate -and $dependsOnRemoval) {
                    # The batch or users to remove first still exist in simulation: Exchange cannot check the creation
                    # before the removal. The CSV was built from the pre-flight; the removals themselves were checked above.
                    Write-EmmLog 'CHANGE' "[WhatIf] New-MigrationBatch -Name '$name' -Local (not sent: the batch or users to remove first still exist in simulation)"
                    $batchDetail += ' - creation not checked by Exchange in simulation (the batch or users to remove first still exist)'
                } else {
                    [void](Invoke-EmmChange -Command 'New-MigrationBatch' -Parameters $p -Simulate:$Simulate)
                    if (-not $Simulate) { [void](Invoke-EmmChange -Command 'Start-MigrationBatch' -Parameters @{ Identity = $name }) }
                }
                Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'New-MigrationBatch -Local + Start' -Status $okStatus -Detail $batchDetail
                foreach ($i in $mailboxItems) {
                    $to = @($(if (Test-EmmMoveType $i.MoveType 'Primary') { $i.Row.TargetDatabase }), $(if (Test-EmmMoveType $i.MoveType 'Archive') { 'archive ' + $i.Row.TargetArchiveDatabase }) | Where-Object { $_ }) -join ', '
                    Add-EmmAction -Workload 'User' -Batch $name -Target $i.Row.PrimarySmtpAddress -Action "Queued in $name $($script:Arrow) $to" -Status $okStatus -Detail $(if ($i.Reason) { $i.Reason } else { $i.MoveType }) -Quiet
                }
            } catch {
                $err = $_.Exception.Message
                Add-EmmAction -Workload 'User' -Batch $name -Target $name -Action 'New-MigrationBatch -Local + Start' -Status Failed -Detail $err
                foreach ($i in $mailboxItems) { Add-EmmAction -Workload 'User' -Batch $name -Target $i.Row.PrimarySmtpAddress -Action "Not submitted" -Status Failed -Detail "Batch $name not created" -Quiet }
            }
        }
        foreach ($i in $folderItems) {
            $p = Get-EmmMoveParameter -Settings $Settings
            $p['Identity'] = $(if ($i.Row.DistinguishedName) { $i.Row.DistinguishedName } else { $i.Row.Guid })
            $p['TargetDatabase'] = $i.Row.TargetDatabase; $p['BatchName'] = $name; $p['SuspendWhenReadyToComplete'] = $true
            try {
                [void](Invoke-EmmChange -Command 'New-MoveRequest' -Parameters $p -Simulate:$Simulate)
                Add-EmmAction -Workload 'User' -Batch $name -Target $i.Row.Name -Action "New-MoveRequest (public folder) $($script:Arrow) $($i.Row.TargetDatabase)" -Status $okStatus -Detail 'suspended when ready to complete'
            } catch {
                Add-EmmAction -Workload 'User' -Batch $name -Target $i.Row.Name -Action "New-MoveRequest (public folder) $($script:Arrow) $($i.Row.TargetDatabase)" -Status Failed -Detail $_.Exception.Message
            }
        }
    }
}

function Wait-EmmSystemMove {
    <# Waits for the system move requests (System.BatchName) to finish, with a progress bar, then records each result. #>
    param([Parameter(Mandatory)]$Settings)
    $name = $Settings.System.BatchName
    $deadline = (Get-Date).AddMinutes($Settings.System.WaitTimeoutMinutes)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    Write-EmmItem Info ("Waiting for the system moves (at most {0} min, Ctrl+C stops waiting, not the moves)..." -f $Settings.System.WaitTimeoutMinutes) -Icon Clock
    $stats = @()
    do {
        $stats = @(Get-MoveRequest -BatchName $name -ResultSize Unlimited -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Get-MoveRequestStatistics -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        $pending = @($stats | Where-Object { [string]$_.Status -notin $script:FinishedMoveStatuses })
        $avg = if ($stats.Count) { [int](@($stats | ForEach-Object { [int](Get-EmmProp $_ 'PercentComplete' 0) } | Measure-Object -Average)[0].Average) } else { 100 }
        Write-Progress -Activity 'System mailbox moves' -Status ('{0} of {1} finished - {2}% - {3}' -f ($stats.Count - $pending.Count), $stats.Count, $avg, (Format-EmmDuration $clock.Elapsed.TotalSeconds)) -PercentComplete ([Math]::Min(100, $avg))
        if (-not $pending.Count) { break }
        Start-Sleep -Seconds $Settings.System.PollSeconds
    } while ((Get-Date) -lt $deadline)
    Write-Progress -Activity 'System mailbox moves' -Completed
    foreach ($s in $stats) {
        $st = [string]$s.Status
        $label = [string](Get-EmmProp $s 'DisplayName' (Get-EmmProp $s 'Alias' ''))
        $status = if ($st -in 'Completed', 'CompletedWithWarning') { 'Success' } elseif ($st -eq 'Failed') { 'Failed' } else { 'Skipped' }
        $detail = if ($status -eq 'Skipped') { "still $st after $($Settings.System.WaitTimeoutMinutes) min: follow it with -Mode Status" } else { "$st $($script:Dot) $([string](Get-EmmProp $s 'TargetDatabase' ''))" }
        if ($st -eq 'Failed') { $detail += ' ' + [string](Get-EmmProp $s 'Message' '') }
        Add-EmmAction -Workload 'System' -Batch $name -Target $label -Action 'Move result' -Status $status -Detail $detail -Quiet:($status -eq 'Success')
    }
    Write-EmmItem $(if (@($stats | Where-Object { [string]$_.Status -notin 'Completed', 'CompletedWithWarning' }).Count) { 'Warn' } else { 'Ok' }) ('System moves: {0} completed of {1} in {2}' -f @($stats | Where-Object { [string]$_.Status -in 'Completed', 'CompletedWithWarning' }).Count, $stats.Count, (Format-EmmDuration $clock.Elapsed.TotalSeconds))
}

function Select-EmmBatchName {
    <#
    .SYNOPSIS
        Turns -Batch values (All, numbers, names, System) into batch names of this tool.
        '3', '03' and 'Batch03' give 'Batch03'; 'System' gives System.BatchName.
    #>
    param([Parameter(Mandatory)]$Settings, [string[]]$Batch, [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Existing)
    $values = @($Batch | Where-Object { $_ } | ForEach-Object { $_ -split '[,\s]+' } | Where-Object { $_ })
    if ($values -contains 'All') { return @($Existing | Sort-Object) }
    $prefix = $Settings.Plan.BatchNamePrefix
    $out = foreach ($v in $values) {
        if ($v -eq 'System' -or $v -eq $Settings.System.BatchName) { $Settings.System.BatchName }
        elseif ($v -match '^\d{1,3}$') { $prefix + ([int]$v).ToString('00') }
        else { $v }
    }
    return @($out | Sort-Object -Unique)
}

function Get-EmmBatchMoveRequest {
    <# Move requests of one tool batch: 'MigrationService:<name>' (mailboxes of the batch) and '<name>' (public folders, system). #>
    param([Parameter(Mandatory)][string]$Name)
    $out = @()
    foreach ($label in @("MigrationService:$Name", $Name)) {
        $out += @(Get-MoveRequest -BatchName $label -ResultSize Unlimited -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    }
    return $out
}

function Complete-EmmBatch {
    <#
    .SYNOPSIS
        Completes migration batches now, or schedules their completion (-CompleteAfter).
    .DESCRIPTION
        Now      : Complete-MigrationBatch (the batch must be Synced), then Resume-MoveRequest for the
                   public folder mailboxes of the batch. 'System' resumes system moves left suspended.
        Scheduled: Set-MoveRequest -CompleteAfter + Resume-MoveRequest on every move request of the
                   batch. Measured on Exchange 2019: Complete-MigrationBatch ignores a completion time
                   and finalises at once, so the move requests are scheduled one by one. The batch then
                   stays 'Synced' in the admin center although its moves are completed: -Mode Status
                   shows the real state, -Mode Cleanup removes such batches.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string[]]$BatchName, [Nullable[datetime]]$CompleteAfter, [switch]$Simulate)
    $okStatus = if ($Simulate) { 'Simulated' } else { 'Success' }
    $scheduled = $null -ne $CompleteAfter -and $CompleteAfter -gt (Get-Date)
    $objects = Get-EmmMigrationObjects -Settings $Settings
    $batches = @{}
    foreach ($b in $objects.ToolBatches) { $batches[$b.Name] = $b }
    $labels = @(Get-EmmManagedName -Objects $objects)
    $resume = {
        # Resumes the move requests of $name waiting for completion (public folders, system, batch-less labels).
        param([object[]]$Requests, [string]$What)
        foreach ($mr in $Requests) {
            $id = [string](Get-EmmProp $mr 'Identity' '')
            try { [void](Invoke-EmmChange -Command 'Resume-MoveRequest' -Parameters @{ Identity = $id } -Simulate:$Simulate); Add-EmmAction -Batch $name -Target $id -Action "Resume-MoveRequest ($What)" -Status $okStatus }
            catch { Add-EmmAction -Batch $name -Target $id -Action "Resume-MoveRequest ($What)" -Status Failed -Detail $_.Exception.Message }
        }
    }
    foreach ($name in $BatchName) {
        if ($name -notin $labels) { Add-EmmAction -Batch $name -Target $name -Action 'Complete' -Status Skipped -Detail 'No migration batch or move request of this tool with this name'; continue }
        $hasBatch = $batches.ContainsKey($name)
        $status = if ($hasBatch) { $batches[$name].Status } else { 'move requests' }
        $moves = @(Get-EmmBatchMoveRequest -Name $name)
        # Move requests labelled with the name itself: public folder mailboxes of a batch, system mailboxes.
        $labelled = @($moves | Where-Object { [string](Get-EmmProp $_ 'BatchName' '') -eq $name })
        $what = if ($name -eq $Settings.System.BatchName) { 'system' } else { 'public folder' }
        if ($scheduled) {
            $scope = $moves
            if ($hasBatch -and $status -in 'Completed', 'CompletedWithErrors') { $scope = $labelled }   # only its late public folder moves are left
            elseif ($hasBatch -and $status -notin 'Synced', 'Syncing') { Add-EmmAction -Batch $name -Target $name -Action 'Schedule completion' -Status Skipped -Detail "Status $status`: only a Synced or Syncing batch can be scheduled"; continue }
            $open = @($scope | Where-Object { [string]$_.Status -notin 'Completed', 'CompletedWithWarning', 'CompletionInProgress', 'Failed' })
            if (-not $open.Count) { Add-EmmAction -Batch $name -Target $name -Action 'Schedule completion' -Status AlreadyDone -Detail 'No move request left to complete'; continue }
            $done = 0; $failed = 0
            foreach ($mr in $open) {
                $id = [string](Get-EmmProp $mr 'Identity' '')
                try {
                    [void](Invoke-EmmChange -Command 'Set-MoveRequest' -Parameters @{ Identity = $id; CompleteAfter = $CompleteAfter } -Simulate:$Simulate)
                    [void](Invoke-EmmChange -Command 'Resume-MoveRequest' -Parameters @{ Identity = $id } -Simulate:$Simulate)
                    $done++
                } catch { $failed++; Add-EmmAction -Batch $name -Target $id -Action 'Schedule completion' -Status Failed -Detail $_.Exception.Message }
            }
            Add-EmmAction -Batch $name -Target $name -Action ("Completion scheduled at {0}" -f ([datetime]$CompleteAfter).ToString('yyyy-MM-dd HH:mm', $script:Invariant)) -Status $(if ($failed) { 'Failed' } else { $okStatus }) -Detail ("{0} move request(s){1}" -f $done, $(if ($hasBatch) { ', the batch stays Synced in the admin center' } else { '' }))
            continue
        }
        $waiting = @($labelled | Where-Object { [string]$_.Status -in 'AutoSuspended', 'Synced' })
        $notReady = {
            # Labelled move requests (public folders) still synchronising: they need another -Mode Complete later.
            $later = @($labelled | Where-Object { [string]$_.Status -notin 'AutoSuspended', 'Synced', 'Completed', 'CompletedWithWarning', 'CompletionInProgress', 'Failed' }).Count
            if ($later) { Add-EmmAction -Batch $name -Target $name -Action "Resume-MoveRequest ($what)" -Status Skipped -Detail "$later $what move request(s) not synchronised yet: run -Mode Complete -Batch $name again when -Mode Status shows them ready" }
        }
        if (-not $hasBatch) {
            if (-not $waiting.Count) { Add-EmmAction -Batch $name -Target $name -Action 'Resume-MoveRequest' -Status Skipped -Detail $(if ($name -eq $Settings.System.BatchName) { 'No system move request waiting for completion (they complete automatically)' } else { 'No move request waiting for completion yet' }); continue }
            & $resume $waiting $what
            & $notReady
            continue
        }
        if ($status -in 'Completed', 'CompletedWithErrors') {
            # The mailboxes of the batch are done; its public folder moves may have synchronised after the batch completion.
            Add-EmmAction -Batch $name -Target $name -Action 'Complete-MigrationBatch' -Status AlreadyDone -Detail $status
            & $resume $waiting $what
            & $notReady
            continue
        }
        if ($status -ne 'Synced') {
            $syncedMoves = @($moves | Where-Object { [string]$_.Status -in 'Synced', 'AutoSuspended' }).Count
            $hint = if ($syncedMoves -and $syncedMoves -eq $moves.Count) { ' - all its move requests are already synced: the batch status is refreshed by the migration service with a delay, try again in a few minutes (guide, Annex A)' } else { '' }
            Add-EmmAction -Batch $name -Target $name -Action 'Complete-MigrationBatch' -Status Skipped -Detail ("Status $status`: Exchange completes only a Synced batch$hint")
            continue
        }
        try {
            [void](Invoke-EmmChange -Command 'Complete-MigrationBatch' -Parameters @{ Identity = $name } -Simulate:$Simulate)
            Add-EmmAction -Batch $name -Target $name -Action 'Complete-MigrationBatch' -Status $okStatus -Detail ("{0} mailbox(es)" -f $batches[$name].Total)
        } catch {
            Add-EmmAction -Batch $name -Target $name -Action 'Complete-MigrationBatch' -Status Failed -Detail $_.Exception.Message
        }
        & $resume $waiting $what
        & $notReady
    }
}

function Test-EmmBatchRemovable {
    <#
    .SYNOPSIS
        Can this batch of the tool be removed without losing anything? One rule for -Mode Cleanup and for
        -Mode Start (a planned mailbox still held by an earlier batch).
    .DESCRIPTION
        Never while one of its moves (public folders included) is not finished: that would cancel synchronised
        moves, or a completion scheduled with -CompleteAfter.
        Removable when finished: Completed / CompletedWithErrors, or Synced with every move finished (the
        on-premises limit: Set-MoveRequest -CompleteAfter completes the moves but the batch stays Synced).
        Failed moves are kept for analysis, unless Cleanup.IncludeFailed, or unless they are the mailboxes
        about to be moved again (-RetriedGuid). A Failed / Stopped / Corrupted batch: same condition.
    #>
    param([Parameter(Mandatory)]$Batch, [AllowEmptyCollection()][object[]]$Moves = @(), [Parameter(Mandatory)]$Settings, [hashtable]$RetriedGuid = @{})
    $pending = @($Moves | Where-Object { $_.Status -notin 'Completed', 'CompletedWithWarning', 'Failed' }).Count
    $failedOther = @($Moves | Where-Object { $_.Status -eq 'Failed' -and -not ($_.ExchangeGuid -and $RetriedGuid.ContainsKey($_.ExchangeGuid)) }).Count
    $retried = $RetriedGuid.Count -gt 0
    $no = { param($Text) [pscustomobject]@{ Removable = $false; Reason = $Text } }
    $yes = { param($Text) [pscustomobject]@{ Removable = $true; Reason = $Text } }
    if ($pending) { return (& $no ('Status {0}, {1} move request(s) not finished' -f $Batch.Status, $pending)) }
    if ($failedOther -and -not $Settings.Cleanup.IncludeFailed) { return (& $no ('Status {0}, {1} failed move(s) kept for analysis (Cleanup.IncludeFailed or -IncludeFailed removes them)' -f $Batch.Status, $failedOther)) }
    if ($Batch.Status -in 'Completed', 'CompletedWithErrors') { return (& $yes $Batch.Status) }
    if ($Batch.Status -eq 'Synced' -and $Moves.Count) { return (& $yes 'Synced, but all its moves are finished (completion scheduled with -CompleteAfter)') }
    if ($Batch.Status -in 'Failed', 'Stopped', 'Corrupted' -and ($Settings.Cleanup.IncludeFailed -or $retried)) { return (& $yes $Batch.Status) }
    return (& $no ('Status {0}' -f $Batch.Status))
}

function Get-EmmCleanupPlan {
    <#
    .SYNOPSIS
        What -Mode Cleanup would remove: computed first, shown to the administrator, then executed
        by Invoke-EmmCleanup. Objects of other tools or of other administrators are never selected.
    .DESCRIPTION
        Migration batches : Completed / CompletedWithErrors (+ Failed, Stopped, Corrupted with
                            Cleanup.IncludeFailed). A batch that still holds failed moves is kept unless
                            Cleanup.IncludeFailed: they are the evidence of what must be retried.
                            A Synced batch is removed only when ALL its move requests are completed
                            (scheduled completion); otherwise it is kept: its synchronised moves would be lost.
        Move requests     : of the tool batches and of System.BatchName, Completed / CompletedWithWarning
                            (+ Failed with Cleanup.IncludeFailed).
        Migration users   : left without their batch (Cleanup.RemoveOrphanMigrationUsers), at execution.
    #>
    param([Parameter(Mandatory)]$Settings, [string[]]$BatchName)
    $objects = Get-EmmMigrationObjects -Settings $Settings
    $selected = if ($BatchName) { @($BatchName) } else { @(Get-EmmManagedName -Objects $objects) }
    $moveOk = @('Completed', 'CompletedWithWarning')
    if ($Settings.Cleanup.IncludeFailed) { $moveOk += @('Failed') }
    $remove = [System.Collections.Generic.List[object]]::new()
    $kept = [System.Collections.Generic.List[object]]::new()
    foreach ($b in @($objects.ToolBatches | Where-Object { $_.Name -in $selected } | Sort-Object Name)) {
        # Every move of the batch: its mailboxes (MigrationService:<name>) and its public folder mailboxes (<name>).
        $test = Test-EmmBatchRemovable -Batch $b -Moves @($objects.ToolMoves | Where-Object { $_.ToolBatch -eq $b.Name }) -Settings $Settings
        if ($test.Removable) { $remove.Add([pscustomobject]@{ Batch = $b; Reason = $test.Reason }) } else { $kept.Add([pscustomobject]@{ Batch = $b; Reason = $test.Reason }) }
    }
    return [pscustomobject]@{
        Batches = $remove.ToArray(); KeptBatches = $kept.ToArray()
        MoveRequests = @($objects.ToolMoves | Where-Object { $_.ToolBatch -in $selected -and $_.Status -in $moveOk })
        KeptMoves = @($objects.ToolMoves | Where-Object { $_.ToolBatch -in $selected -and $_.Status -notin $moveOk }).Count
        Selected = $selected; Filtered = [bool]$BatchName
    }
}

function Invoke-EmmCleanup {
    <#
    .SYNOPSIS
        Executes a cleanup plan (Get-EmmCleanupPlan): move requests first, then batches, then the
        migration users left without their batch.
    #>
    param([Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$CleanupPlan, [switch]$Simulate)
    $okStatus = if ($Simulate) { 'Simulated' } else { 'Success' }
    foreach ($k in $CleanupPlan.KeptBatches) { Add-EmmAction -Batch $k.Batch.Name -Target $k.Batch.Name -Action 'Kept' -Status Skipped -Detail $k.Reason }
    foreach ($g in @($CleanupPlan.MoveRequests | Group-Object ToolBatch | Sort-Object Name)) {
        $ok = 0; $ko = 0
        foreach ($mr in $g.Group) {
            $label = if ($mr.DisplayName) { $mr.DisplayName } else { $mr.Identity }
            try { [void](Invoke-EmmChange -Command 'Remove-MoveRequest' -Parameters @{ Identity = $mr.Identity } -Simulate:$Simulate); $ok++; Add-EmmAction -Batch $g.Name -Target $label -Action 'Remove-MoveRequest' -Status $okStatus -Detail $mr.Status -Quiet }
            catch { $ko++; Add-EmmAction -Batch $g.Name -Target $label -Action 'Remove-MoveRequest' -Status Failed -Detail $_.Exception.Message }
        }
        Write-EmmItem $(if ($ko) { 'Warn' } else { 'Ok' }) ('{0,-22} {1} move request(s) removed{2}' -f $g.Name, $ok, $(if ($ko) { ", $ko failed" } else { '' })) -Icon $(if ($ko) { '' } else { 'Broom' })
    }
    foreach ($item in $CleanupPlan.Batches) {
        $b = $item.Batch
        try { [void](Invoke-EmmChange -Command 'Remove-MigrationBatch' -Parameters @{ Identity = $b.Name; Force = $true } -Simulate:$Simulate); Add-EmmAction -Batch $b.Name -Target $b.Name -Action 'Remove-MigrationBatch' -Status $okStatus -Detail ('{0}, {1} mailbox(es)' -f $item.Reason, $b.Total) }
        catch { Add-EmmAction -Batch $b.Name -Target $b.Name -Action 'Remove-MigrationBatch' -Status Failed -Detail $_.Exception.Message }
    }
    $orphanCount = 0
    if ($Settings.Cleanup.RemoveOrphanMigrationUsers) {
        # The migration service needs a few seconds to forget the users of a removed batch.
        if ($CleanupPlan.Batches.Count -and -not $Simulate -and $Settings.Cleanup.SettleSeconds) { Start-Sleep -Seconds $Settings.Cleanup.SettleSeconds }
        # Fail closed: without the list of batches, every migration user would look orphan.
        try { $current = @(Get-MigrationBatch -ErrorAction Stop -WarningAction SilentlyContinue) }
        catch { Write-EmmItem Warn "Orphan migration users not checked: the migration batches cannot be read ($($_.Exception.Message))."; return $orphanCount }
        $existing = @{}
        foreach ($b in $current) { $existing[[string](Get-EmmProp $b 'Identity' (Get-EmmProp $b 'Name' ''))] = $true }
        if ($Simulate) { foreach ($item in $CleanupPlan.Batches) { $existing.Remove($item.Batch.Name) } }
        # Only users of a batch of this tool that no longer exists, and of the selected batches when -Batch is used.
        $orphans = @(Get-MigrationUser -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue | Where-Object {
                $batch = [string](Get-EmmProp $_ 'BatchId' '')
                $batch -and (Get-EmmToolBatchName $batch $Settings) -and -not $existing.ContainsKey($batch) -and (-not $CleanupPlan.Filtered -or $batch -in $CleanupPlan.Selected)
            })
        foreach ($u in $orphans) {
            $id = [string](Get-EmmProp $u 'Identity' '')
            try { [void](Invoke-EmmChange -Command 'Remove-MigrationUser' -Parameters @{ Identity = $id } -Simulate:$Simulate); $orphanCount++; Add-EmmAction -Batch ([string](Get-EmmProp $u 'BatchId' '')) -Target $id -Action 'Remove-MigrationUser (orphan)' -Status $okStatus -Quiet }
            catch { Add-EmmAction -Target $id -Action 'Remove-MigrationUser (orphan)' -Status Failed -Detail $_.Exception.Message }
        }
        if ($orphanCount) { Write-EmmItem Ok ('{0} orphan migration user(s) removed' -f $orphanCount) -Icon Broom }
    }
    return $orphanCount
}

#endregion
#region 7. Status ---------------------------------------------------------------------------------------------

function Get-EmmStatusGroup {
    <# Display group of a move or migration user status: Completed, Ready (synced, waiting for completion), InProgress, Failed. #>
    param([string]$Status)
    switch -Regex ($Status) {
        '^(Completed|CompletedWithWarning|CompletedWithErrors|Finalized)$' { return 'Completed' }
        '^(Synced|AutoSuspended)$' { return 'Ready' }
        '^(Failed|Suspended|Stopped|Corrupted|IncrementalSyncFailed|CompletionFailed|SyncFailed)$' { return 'Failed' }
        default { return 'InProgress' }
    }
}

function Get-EmmStatus {
    <#
    .SYNOPSIS
        Real state of the migration: every tool batch with its move requests, and the system moves.
    .DESCRIPTION
        The move request statistics (mailbox replication service) are the reference: status, percent,
        size, target, StatusDetail (a 'StalledDueTo...' detail = the move waits: quarantined mailbox,
        busy target...). The migration users add the batch view and the users that have no move
        request yet (validation errors).
    #>
    param([Parameter(Mandatory)]$Settings, [string[]]$BatchName)
    $objects = Get-EmmMigrationObjects -Settings $Settings
    $names = @(Get-EmmManagedName -Objects $objects)
    $systemName = $Settings.System.BatchName
    if ($BatchName) { $names = @($names | Where-Object { $_ -in $BatchName }) }
    $byName = @{}; foreach ($b in $objects.ToolBatches) { $byName[$b.Name] = $b }
    $rows = [System.Collections.Generic.List[object]]::new()
    $summaries = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($name in @($names | Sort-Object @{ Expression = { if ($_ -eq $systemName) { 0 } else { 1 } } }, { $_ })) {
        $index++
        Write-Progress -Activity 'Reading the migration state' -Status $name -PercentComplete ([int](100 * ($index - 1) / [Math]::Max(1, $names.Count)))
        $isSystem = $name -eq $systemName
        $hasBatch = $byName.ContainsKey($name)
        $batchStatus = if ($hasBatch) { $byName[$name].Status } else { 'Move requests' }
        $stats = @(Get-EmmBatchMoveRequest -Name $name | Get-MoveRequestStatistics -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        $users = @(if ($hasBatch) { Get-MigrationUser -BatchId $name -ResultSize Unlimited -ErrorAction SilentlyContinue -WarningAction SilentlyContinue })
        $userByGuid = @{}; $userByAlias = @{}
        foreach ($u in $users) {
            $g = ([string](Get-EmmProp $u 'MailboxGuid' '')).ToLowerInvariant()
            $mail = ([string](Get-EmmProp $u 'MailboxEmailAddress' (Get-EmmProp $u 'Identity' ''))).ToLowerInvariant()
            $o = [pscustomobject]@{ User = $u; Mail = $mail; Used = $false }
            if ($g) { $userByGuid[$g] = $o }
            if ($mail -match '^([^@]+)@') { $userByAlias[$Matches[1]] = $o }
        }
        $batchRows = [System.Collections.Generic.List[object]]::new()
        foreach ($s in $stats) {
            $guid = ([string](Get-EmmProp $s 'ExchangeGuid' '')).ToLowerInvariant()
            $alias = ([string](Get-EmmProp $s 'Alias' '')).ToLowerInvariant()
            $u = if ($guid -and $userByGuid.ContainsKey($guid)) { $userByGuid[$guid] } elseif ($alias -and $userByAlias.ContainsKey($alias)) { $userByAlias[$alias] } else { $null }
            if ($u) { $u.Used = $true }
            $status = [string](Get-EmmProp $s 'Status' '')
            $detail = [string](Get-EmmProp $s 'StatusDetail' '')
            $batchRows.Add([pscustomobject][ordered]@{
                    Workload = $(if ($isSystem) { 'System' } else { 'User' }); Batch = $name; BatchStatus = $batchStatus
                    DisplayName = [string](Get-EmmProp $s 'DisplayName' $alias); Mail = $(if ($u) { $u.Mail } else { '' }); Alias = [string](Get-EmmProp $s 'Alias' '')
                    Status = $status; Group = Get-EmmStatusGroup $status; ServiceStatus = $(if ($u) { [string](Get-EmmProp $u.User 'Status' '') } else { '' })
                    StatusDetail = $detail; Quarantined = [bool]($detail -match '^Stalled|Quarantin')
                    Percent = [int](Get-EmmProp $s 'PercentComplete' 0)
                    SizeMB = [Math]::Round((ConvertTo-EmmMB (Get-EmmProp $s 'TotalMailboxSize')) + (ConvertTo-EmmMB (Get-EmmProp $s 'TotalArchiveSize')), 2)
                    TransferredMB = ConvertTo-EmmMB (Get-EmmProp $s 'BytesTransferred')
                    SourceDatabase = Get-EmmName (Get-EmmProp $s 'SourceDatabase' ''); TargetDatabase = Get-EmmName (Get-EmmProp $s 'TargetDatabase' '')
                    TargetArchiveDatabase = Get-EmmName (Get-EmmProp $s 'TargetArchiveDatabase' '')
                    LastUpdate = Format-EmmDate (Get-EmmProp $s 'LastUpdateTimestamp')
                    Message = [string](Get-EmmProp $s 'Message' ''); FailureType = [string](Get-EmmProp $s 'FailureType' '')
                    ExchangeGuid = $guid
                })
        }
        foreach ($o in @($userByGuid.Values) + @($userByAlias.Values)) {
            if ($o.Used) { continue }
            $o.Used = $true
            $status = [string](Get-EmmProp $o.User 'Status' '')
            $batchRows.Add([pscustomobject][ordered]@{
                    Workload = 'User'; Batch = $name; BatchStatus = $batchStatus; DisplayName = $o.Mail; Mail = $o.Mail; Alias = ''
                    Status = $status; Group = Get-EmmStatusGroup $status; ServiceStatus = $status; StatusDetail = 'No move request yet'; Quarantined = $false
                    Percent = 0; SizeMB = 0.0; TransferredMB = 0.0; SourceDatabase = ''; TargetDatabase = ''; TargetArchiveDatabase = ''
                    LastUpdate = Format-EmmDate (Get-EmmProp $o.User 'LastSuccessfulSyncTime')
                    Message = [string](Get-EmmProp $o.User 'ErrorSummary' (Get-EmmProp $o.User 'Error' '')); FailureType = ''; ExchangeGuid = ''
                })
        }
        foreach ($r in $batchRows) { $rows.Add($r) }
        $count = { param($g) @($batchRows | Where-Object { $_.Group -eq $g }).Count }
        $total = $batchRows.Count
        $completed = & $count 'Completed'; $ready = & $count 'Ready'; $failed = & $count 'Failed'
        $label = if ($isSystem) { 'System' } else { $name }
        $next = if ($total -and $completed -eq $total) { "-Mode Cleanup -Batch $label" }
            elseif ($batchStatus -eq 'Synced' -or ($ready -and $batchStatus -in 'Move requests', 'Completed', 'CompletedWithErrors')) { "-Mode Complete -Batch $label" }
            elseif ($failed) { 'Check the failed mailboxes' }
            elseif (& $count 'InProgress') { 'Synchronising' }
            elseif ($batchStatus -in 'Completed', 'CompletedWithErrors') { "-Mode Cleanup -Batch $label" }
            else { 'Synchronising' }
        $summaries.Add([pscustomobject][ordered]@{
                Name = $name; Workload = $(if ($isSystem) { 'System' } else { 'User' }); Status = $batchStatus; Mailboxes = $total
                Completed = $completed; Ready = $ready; InProgress = & $count 'InProgress'; Failed = $failed
                Quarantined = @($batchRows | Where-Object { $_.Quarantined }).Count
                Percent = $(if ($total) { [int](@($batchRows | Measure-Object Percent -Average)[0].Average) } else { 0 })
                SizeMB = [Math]::Round([double](@($batchRows | Measure-Object SizeMB -Sum)[0].Sum), 2)
                TransferredMB = [Math]::Round([double](@($batchRows | Measure-Object TransferredMB -Sum)[0].Sum), 2)
                Next = $next
            })
    }
    Write-Progress -Activity 'Reading the migration state' -Completed
    return [pscustomobject]@{ Batches = $summaries.ToArray(); Rows = $rows.ToArray(); OtherBatches = $objects.OtherBatches; OtherMoves = $objects.OtherMoves }
}

#endregion
#region 8. Reports -------------------------------------------------------------------------------------------

function New-EmmRunDirectory {
    <# Folder of one execution: <OutputPath>\yyyy-MM-dd_HHmmss_<Mode>[_<Suffix>]. #>
    param([Parameter(Mandatory)][string]$OutputPath, [Parameter(Mandatory)][string]$Mode, [string]$Suffix)
    $name = '{0}_{1}{2}' -f (Get-Date).ToString('yyyy-MM-dd_HHmmss', $script:Invariant), $Mode, $(if ($Suffix) { '_' + $Suffix } else { '' })
    $path = Join-Path $OutputPath $name
    $n = 1
    while (Test-Path -LiteralPath $path) { $n++; $path = Join-Path $OutputPath ('{0}-{1}' -f $name, $n) }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function ConvertTo-EmmJsonString {
    <# JSON string literal. <, > and & are escaped too, so the JSON can sit inside a <script> element. #>
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return 'null' }
    $s = $Text.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
    $s = $s.Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
    if ($s.IndexOfAny($script:JsonControl) -ge 0) {
        $sb = [System.Text.StringBuilder]::new()
        foreach ($ch in $s.ToCharArray()) { if ([int]$ch -lt 32) { [void]$sb.AppendFormat('\u{0:x4}', [int]$ch) } else { [void]$sb.Append($ch) } }
        $s = $sb.ToString()
    }
    return '"' + $s + '"'
}
$script:JsonControl = [char[]](@(0..31) | ForEach-Object { [char]$_ })

function Add-EmmJsonValue {
    param([Parameter(Mandatory)][System.Text.StringBuilder]$Builder, [AllowNull()]$Value)
    if ($null -eq $Value) { [void]$Builder.Append('null'); return }
    if ($Value -is [string] -or $Value -is [char] -or $Value -is [guid] -or $Value -is [enum]) { [void]$Builder.Append((ConvertTo-EmmJsonString ([string]$Value))); return }
    if ($Value -is [bool] -or $Value -is [switch]) { [void]$Builder.Append($(if ($Value) { 'true' } else { 'false' })); return }
    if ($Value -is [datetime]) { [void]$Builder.Append((ConvertTo-EmmJsonString ($Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant)))); return }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        $d = [double]$Value
        [void]$Builder.Append($(if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { 'null' } else { $d.ToString('R', $script:Invariant) }))
        return
    }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte] -or $Value -is [uint32] -or $Value -is [uint64] -or $Value -is [sbyte] -or $Value -is [uint16]) {
        [void]$Builder.Append([Convert]::ToString($Value, $script:Invariant)); return
    }
    if ($Value -is [System.Collections.IDictionary]) {
        [void]$Builder.Append('{'); $first = $true
        foreach ($k in $Value.Keys) {
            if (-not $first) { [void]$Builder.Append(',') }; $first = $false
            [void]$Builder.Append((ConvertTo-EmmJsonString ([string]$k))).Append(':')
            Add-EmmJsonValue -Builder $Builder -Value $Value[$k]
        }
        [void]$Builder.Append('}'); return
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        [void]$Builder.Append('['); $first = $true
        foreach ($item in $Value) {
            if (-not $first) { [void]$Builder.Append(',') }; $first = $false
            Add-EmmJsonValue -Builder $Builder -Value $item
        }
        [void]$Builder.Append(']'); return
    }
    [void]$Builder.Append('{'); $first = $true
    foreach ($p in $Value.PSObject.Properties) {
        if (-not $first) { [void]$Builder.Append(',') }; $first = $false
        [void]$Builder.Append((ConvertTo-EmmJsonString $p.Name)).Append(':')
        Add-EmmJsonValue -Builder $Builder -Value $p.Value
    }
    [void]$Builder.Append('}')
}

function ConvertTo-EmmJson {
    <#
    .SYNOPSIS
        JSON text of a value, built by the tool itself: ConvertTo-Json of Windows PowerShell 5.1 writes dates
        as \/Date(...)\/, stops at depth 2 and does not escape < > &; here the output is invariant-culture
        and safe inside <script>.
    #>
    param([AllowNull()]$Value)
    $sb = [System.Text.StringBuilder]::new()
    Add-EmmJsonValue -Builder $sb -Value $Value
    return $sb.ToString()
}

function Add-EmmJsonTable {
    <# Rows as {"columns":[...],"rows":[[...],...]}: compact, and fast (no function call per cell for common types). #>
    param([Parameter(Mandatory)][System.Text.StringBuilder]$Builder, [AllowEmptyCollection()][object[]]$Rows = @(), [string[]]$Columns)
    if (-not $Columns) { $Columns = if ($Rows.Count) { @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name }) } else { @() } }
    [void]$Builder.Append('{"columns":'); Add-EmmJsonValue -Builder $Builder -Value @($Columns); [void]$Builder.Append(',"rows":[')
    $firstRow = $true
    foreach ($row in $Rows) {
        if (-not $firstRow) { [void]$Builder.Append(',') }; $firstRow = $false
        [void]$Builder.Append('[')
        for ($c = 0; $c -lt $Columns.Count; $c++) {
            if ($c) { [void]$Builder.Append(',') }
            $p = $row.PSObject.Properties[$Columns[$c]]
            $v = if ($p) { $p.Value } else { $null }
            if ($null -eq $v) { [void]$Builder.Append('null') }
            elseif ($v -is [string]) { [void]$Builder.Append((ConvertTo-EmmJsonString $v)) }
            elseif ($v -is [int] -or $v -is [long]) { [void]$Builder.Append([Convert]::ToString($v, $script:Invariant)) }
            elseif ($v -is [double]) { [void]$Builder.Append($(if ([double]::IsNaN($v) -or [double]::IsInfinity($v)) { 'null' } else { $v.ToString('R', $script:Invariant) })) }
            elseif ($v -is [bool]) { [void]$Builder.Append($(if ($v) { 'true' } else { 'false' })) }
            else { Add-EmmJsonValue -Builder $Builder -Value $v }
        }
        [void]$Builder.Append(']')
    }
    [void]$Builder.Append(']}')
}

function Export-EmmCsv {
    <#
    .SYNOPSIS
        CSV file in UTF-8 with BOM (Excel shows accents). With the ';' delimiter (French Excel) the
        decimal separator is a comma. A text starting with = + - @ is prefixed with ' so Excel never
        runs it as a formula.
    #>
    param([AllowEmptyCollection()][object[]]$Rows = @(), [string[]]$Columns, [Parameter(Mandatory)][string]$Path, [string]$Delimiter = ';')
    if (-not $Columns) { $Columns = if ($Rows.Count) { @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name }) } else { @() } }
    $decimal = if ($Delimiter -eq ';') { ',' } else { '.' }
    $quote = [char[]]@($Delimiter[0], '"', "`r", "`n")
    $esc = {
        param($v)
        if ($null -eq $v) { return '' }
        if ($v -is [double] -or $v -is [decimal] -or $v -is [single]) { return ([double]$v).ToString('0.##', $script:Invariant).Replace('.', $decimal) }
        if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant) }
        if ($v -is [array]) { $v = ($v -join ', ') }
        $s = [string]$v
        if ($s.Length -and '=+-@'.IndexOf($s[0]) -ge 0) { $s = "'" + $s }
        if ($s.IndexOfAny($quote) -ge 0) { $s = '"' + $s.Replace('"', '""') + '"' }
        return $s
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine((@($Columns | ForEach-Object { & $esc $_ }) -join $Delimiter))
    foreach ($row in $Rows) {
        $cells = for ($c = 0; $c -lt $Columns.Count; $c++) { $p = $row.PSObject.Properties[$Columns[$c]]; & $esc $(if ($p) { $p.Value } else { $null }) }
        [void]$sb.AppendLine(($cells -join $Delimiter))
    }
    [IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))
}

function New-EmmReport {
    <#
    .SYNOPSIS
        Writes the HTML report of an execution (and its CSV files).
    .DESCRIPTION
        The HTML page is templates\Report.template.html with the marker %%DATA%% replaced by the data
        as JSON: {"kind": ..., "meta": {...}, "tables": {name: {"columns": [...], "rows": [[...]]}}}.
        The page is self-contained (no external resource): it can be sent by e-mail or opened from a share.
    .PARAMETER Tables
        Name -> rows. .PARAMETER Columns  Name -> column list (optional, default = properties of the first row).
    .PARAMETER Csv
        CSV file name -> table name.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Inventory', 'Plan', 'Status', 'Start', 'Complete', 'Cleanup')][string]$Kind,
        [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)][System.Collections.IDictionary]$Meta,
        [System.Collections.IDictionary]$Tables = @{}, [System.Collections.IDictionary]$Columns = @{}, [System.Collections.IDictionary]$Csv = @{},
        [string]$HtmlName, [int]$RefreshSeconds = 0
    )
    [void][IO.Directory]::CreateDirectory($Directory)
    $files = [System.Collections.Generic.List[object]]::new()
    $m = [ordered]@{
        tool = $script:ToolName; toolVersion = $script:ToolVersion; author = 'Nicolas Fabert'
        generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant); organization = $Settings.Report.Organization
        refreshSeconds = $RefreshSeconds; batchPrefix = $Settings.Plan.BatchNamePrefix; systemBatch = $Settings.System.BatchName
    }
    foreach ($k in $Meta.Keys) { $m[$k] = $Meta[$k] }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('{"kind":').Append((ConvertTo-EmmJsonString $Kind)).Append(',"meta":')
    Add-EmmJsonValue -Builder $sb -Value $m
    [void]$sb.Append(',"tables":{')
    $first = $true
    foreach ($name in $Tables.Keys) {
        if (-not $first) { [void]$sb.Append(',') }; $first = $false
        [void]$sb.Append((ConvertTo-EmmJsonString ([string]$name))).Append(':')
        $cols = if ($Columns.Contains($name)) { [string[]]$Columns[$name] } else { $null }
        Add-EmmJsonTable -Builder $sb -Rows @($Tables[$name]) -Columns $cols
    }
    [void]$sb.Append('}}')
    $template = [IO.File]::ReadAllText($Settings.Report.TemplatePath)
    $marker = '%%DATA%%'
    $count = ([regex]::Matches($template, [regex]::Escape($marker))).Count
    if ($count -ne 1) { throw "The report template must contain the marker $marker exactly once (found $count): $($Settings.Report.TemplatePath)" }
    if (-not $HtmlName) { $HtmlName = "$Kind.html" }
    $html = Join-Path $Directory $HtmlName
    [IO.File]::WriteAllText($html, $template.Replace($marker, $sb.ToString()), (New-Object Text.UTF8Encoding($false)))
    $files.Add([pscustomobject]@{ Kind = 'HTML'; Path = $html; Rows = $null; Bytes = (New-Object IO.FileInfo($html)).Length })
    foreach ($fileName in $Csv.Keys) {
        $tableName = $Csv[$fileName]
        $rows = @($Tables[$tableName])
        $path = Join-Path $Directory $fileName
        $cols = if ($Columns.Contains($tableName)) { [string[]]$Columns[$tableName] } else { $null }
        Export-EmmCsv -Rows $rows -Columns $cols -Path $path -Delimiter $Settings.Report.CsvDelimiter
        $files.Add([pscustomobject]@{ Kind = 'CSV'; Path = $path; Rows = $rows.Count; Bytes = (New-Object IO.FileInfo($path)).Length })
    }
    return $files.ToArray()
}

#endregion
