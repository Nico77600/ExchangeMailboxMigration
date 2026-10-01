#Requires -Version 5.1
<#
.SYNOPSIS
    Copies the files needed to run Exchange Mailbox Migration into a separate folder, ready to be zipped
    and copied to an Exchange server.

.DESCRIPTION
    The package contains only what Invoke-ExchangeMailboxMigration.ps1 needs at run time, plus the HTML guide:
        Invoke-ExchangeMailboxMigration.ps1, ExchangeMailboxMigration.psd1, ExchangeMailboxMigration.psm1,
        config\, templates\, docs\ExchangeMailboxMigration-Guide.html, README.md, CHANGELOG.md, LICENSE
    It never copies reports\, logs\ or tests\.

    -ConfigPath replaces the delivered configuration file by the configuration of an organisation (for
    example a copy kept outside the repository), so the package is ready for that organisation.

.PARAMETER Destination
    Package folder. Default: package\ExchangeMailboxMigration-<version>, next to the tool folder.

.PARAMETER ConfigPath
    Configuration file to put in the package instead of the delivered one. It is checked first.

.PARAMETER Force
    Replace the destination folder if it already contains a package. A folder that contains a reports\
    sub-folder (a package that has been used) is never replaced.

.EXAMPLE
    .\tools\New-EmmPackage.ps1
    Creates ..\package\ExchangeMailboxMigration-2.0.0.

.EXAMPLE
    .\tools\New-EmmPackage.ps1 -ConfigPath D:\Private\contoso.config.psd1 -Force
    Package with the configuration of one organisation.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [string]$ConfigPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = (Import-PowerShellDataFile (Join-Path $root 'ExchangeMailboxMigration.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\ExchangeMailboxMigration-$version" }
if (-not [IO.Path]::IsPathRooted($Destination)) { $Destination = Join-Path (Get-Location).Path $Destination }
$Destination = [IO.Path]::GetFullPath($Destination).TrimEnd('\')

$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the tool folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-ExchangeMailboxMigration.ps1'))) { throw "The destination is not an Exchange Mailbox Migration package, it is not replaced: $Destination" }
    if (Test-Path -LiteralPath (Join-Path $Destination 'reports')) { throw "The destination contains a reports folder (a package in use), it is not replaced: $Destination" }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

# ---- Files needed at run time ---------------------------------------------------------------------------
$files = 'Invoke-ExchangeMailboxMigration.ps1', 'ExchangeMailboxMigration.psd1', 'ExchangeMailboxMigration.psm1', 'templates\Report.template.html',
    'config\ExchangeMailboxMigration.config.psd1', 'docs\ExchangeMailboxMigration-Guide.html', 'README.md', 'CHANGELOG.md', 'LICENSE'
foreach ($f in $files) {
    $source = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Configuration of an organisation (optional), checked with the module itself --------------------------
if ($ConfigPath) {
    Import-Module (Join-Path $root 'ExchangeMailboxMigration.psd1') -Force
    [void](Import-EmmConfiguration -Path $ConfigPath -Root $root)
    Copy-Item -LiteralPath $ConfigPath -Destination (Join-Path $Destination 'config\ExchangeMailboxMigration.config.psd1') -Force
}

# ---- Checks ----------------------------------------------------------------------------------------------
$problems = [System.Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests') { if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") } }
# Filter on the extension: Windows PowerShell 5.1 ignores -Include together with -LiteralPath.
foreach ($f in @(Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' })) {
    $errors = $null; [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errors)
    if (@($errors).Count) { $problems.Add("Syntax error in $($f.Name): $($errors[0].Message)") }
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

$all = @(Get-ChildItem -LiteralPath $Destination -Recurse -File)
Write-Host ''
Write-Host "  Exchange Mailbox Migration $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host ("  Config   : {0}" -f $(if ($ConfigPath) { "from $ConfigPath (checked)" } else { 'delivered example - adapt the Databases section (guide, chapter 7)' }))
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,12:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
