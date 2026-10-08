#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the graphics of the GitHub README from the Exchange Mailbox Migration guide, in a light and a
    dark version: banner, what it does and the safety rules, how it works.

.DESCRIPTION
    GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are lost. This
    tool renders them as images with the CSS of the built HTML guide and the icons of
    tools\Build-Documentation.ps1, so that the README and the guide always look the same. The README shows
    them with <picture>, which picks the light or dark image from the theme of the reader.

    Sources:
      package\docs\ExchangeMailboxMigration-Guide.md     the cards and flow blocks (chapters 2 and 4)
      package\docs\ExchangeMailboxMigration-Guide.html   the CSS (run tools\Build-Documentation.ps1 first)
      package\ExchangeMailboxMigration.psd1              the version of the badge
      tools\Build-Documentation.ps1                      the icons

    Screenshots: Microsoft Edge in headless mode, with a temporary profile, 2x resolution. Only local files
    are opened. Output: package\docs\images\readme-<name>-light.png and readme-<name>-dark.png. The package
    tool never copies them: they belong to the GitHub page only.

.PARAMETER OutputFolder
    Default: package\docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (the HTML pages of the graphics) and shows its path.

.EXAMPLE
    .\tools\Build-Documentation.ps1; .\tools\New-ReadmeImages.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0 - adapted from tools\New-ReadmeImages.ps1 of X-TAP Sharing Migration (version 1.0.3)
    Part of : Exchange Mailbox Migration (repository tool, not in the package)
    Workstation tool: PowerShell 7.4+, never connected to Exchange. The migration tool itself runs in
    Windows PowerShell 5.1 only.
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'package\docs\images' }

#region Assets of the guide ------------------------------------------------------------------------
function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path $Root 'package\docs\ExchangeMailboxMigration-Guide.html'
    $guideMd = Join-Path $Root 'package\docs\ExchangeMailboxMigration-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'package\docs\ExchangeMailboxMigration-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = (Import-PowerShellDataFile (Join-Path $Root 'package\ExchangeMailboxMigration.psd1')).ModuleVersion
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines) {
    # The pipeline of the guide, on one line: a node per mode, an arrow with its label between them.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    "<div class=""flow rb-flow"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}
#endregion

#region Styles of the graphics, on top of the CSS of the guide -------------------------------------
$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 270px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-cards3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.cards code { white-space: nowrap; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px 14px; gap: 2px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 8px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-text { font-size: 11.5px; }
.rb-flow .flow-arrow { min-width: 0; width: 72px; flex: 0 0 72px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 72px; }
.rb-space { height: 18px; }
'@
#endregion

#region Rendering (Microsoft Edge, headless) -------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    $profilePath = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profilePath`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, [Text.UTF8Encoding]::new($false))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}
#endregion

#region Main ---------------------------------------------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('emm-readme-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$Script:assets = Get-ReadmeAssets -Root $root
if ($assets.Flows.Count -lt 1 -or $assets.Cards.Count -lt 3) { throw 'The guide must hold a flow block (how it works) and 3 cards blocks (what it does, the two workloads, the safety rules).' }
$mid = '&middot;'

try {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'

    # Banner: the hero of the guide, with the key figures of chapters 2, 3 and 4.
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')Windows PowerShell 5.1 only</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'database' 'icon-sm')Exchange Server 2016 / 2019 / SE</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'people' 'icon-sm')User, shared, resource, public folder mailboxes and archives</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'mail')</div><div><div class=""eyebrow"">Exchange Server $mid Mailbox moves $mid PowerShell</div><h1>Exchange Mailbox Migration</h1></div></div>" +
        "<p class=""lead"">Moves the mailboxes of an Exchange Server organisation <strong>from source databases to target databases</strong> with migration batches &mdash; <strong>system and user mailboxes kept apart</strong>, <strong>monitoring mailboxes never moved</strong> &mdash; with an <strong>HTML and CSV report</strong> at every step.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>6</b><span><strong>modes, one per step</strong>Inventory, Plan, Start, Status, Complete, Cleanup</span></div>" +
        "<div class=""rb-stat""><b>2</b><span><strong>workloads kept apart</strong>system mailboxes, then user mailboxes</span></div>" +
        "<div class=""rb-stat""><b>0</b><span><strong>monitoring mailboxes</strong>never moved, whatever the configuration</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1120

    # Why: what the tool does (first cards block of the guide) and the safety rules (chapter 4).
    $why = "<div class=""rb-caption"">What the tool does <span>$mid and what it never does</span></div>" +
        (ConvertTo-ReadmeCards $assets.Cards[0].Lines 'rb-cards2') + "<div class=""rb-space""></div>" +
        "<div class=""rb-caption"">Safety rules <span>$mid checked at every step, nothing else is touched</span></div>" +
        (ConvertTo-ReadmeCards $assets.Cards[2].Lines 'rb-cards3')
    New-ReadmeGraphic -Name 'principles' -Body $why -Width 1120

    # How it works: the flow of chapter 2, then the two workloads (second cards block).
    $howItWorks = "<div class=""rb-caption"">One script, one mode per step <span>$mid read-only until Start</span></div>" +
        (ConvertTo-ReadmeFlow $assets.Flows[0].Lines) + "<div class=""rb-space""></div>" +
        "<div class=""rb-caption"">Two workloads <span>$mid the plan keeps them apart</span></div>" +
        (ConvertTo-ReadmeCards $assets.Cards[1].Lines 'rb-cards2')
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1120

    Get-ChildItem $OutputFolder -Filter 'readme-*.png' | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
} finally {
    # Edge helper processes of the temporary profile, if any are left.
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion
