<#
.SYNOPSIS
    Install onenote-plugin into a DSH profile.

.DESCRIPTION
    1. Sync this directory to <DSH_HOME>\plugins-src\onenote-plugin (staging)
    2. Copy it into <profile>\node_modules\onenote-plugin
    3. In the profile package.json: register the file: dependency and the
       dsh.profile.bundles entry for onenote-plugin, and remove the old
       dsh-onenote dependency/bundle when present (backup first)
    4. If the dsh CLI is available, run pnpm so pnpm-lock.yaml stays in sync
    5. Print restart instructions

    -DryRun prints the resolved paths and the planned edits, then exits
    without writing anything.

    Machine-specific section paths stay OUT of the repo: the shipped
    cordis.patch.yml has extraSections: []; -ExtraSectionPath fills it into the
    installed copy only.

    Restart the DSH desktop app after installing: bundles are composed at
    startup.
#>
[CmdletBinding()]
param(
    [string]$ProfileName = 'desktop',
    [string]$Registry = '',
    [string]$DshCli = '',
    [string]$ExtraSectionName = '',
    [string]$ExtraSectionPath = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$source = $PSScriptRoot
# DSH_HOME wins: when it is set, the desktop app boots the profile under it and
# a profile under %USERPROFILE%\.dsh is never loaded (that mismatch silently
# installs into the wrong home). Fall back to the legacy default.
$dshHome = $env:DSH_HOME
if ([string]::IsNullOrWhiteSpace($dshHome)) { $dshHome = Join-Path $env:USERPROFILE '.dsh' }
$dshHome = $dshHome.TrimEnd('\', '/')
$profileDir = Join-Path $dshHome "profiles\$ProfileName"
$staging = Join-Path $dshHome 'plugins-src\onenote-plugin'
$target = Join-Path $profileDir 'node_modules\onenote-plugin'
$pkgPath = Join-Path $profileDir 'package.json'
$packageName = 'onenote-plugin'
$legacyName = 'dsh-onenote'

if (-not (Test-Path -LiteralPath $pkgPath)) {
    throw "profile not found: $profileDir (expected $pkgPath)"
}

$depValue = 'file:' + ($staging -replace '\\', '/')
$utf8 = New-Object System.Text.UTF8Encoding($false)
$pkg = [IO.File]::ReadAllText($pkgPath, $utf8) | ConvertFrom-Json

$hasDeps = $null -ne $pkg.PSObject.Properties['dependencies']
$hasLegacyDep = $hasDeps -and ($pkg.dependencies.PSObject.Properties.Name -contains $legacyName)
$bundlesBefore = @()
if ($null -ne $pkg.PSObject.Properties['dsh'] -and $null -ne $pkg.dsh.PSObject.Properties['profile']) {
    $bundlesBefore = @($pkg.dsh.profile.bundles)
}
$hasLegacyBundle = $bundlesBefore -contains $legacyName

Write-Host "DSH home  : $dshHome"
Write-Host "profile   : $profileDir"
Write-Host "staging   : $staging"
Write-Host "target    : $target"
Write-Host "dependency: $depValue"

if ($DryRun) {
    $sectionName = $ExtraSectionName
    if ([string]::IsNullOrWhiteSpace($sectionName)) {
        $sectionName = [IO.Path]::GetFileNameWithoutExtension($ExtraSectionPath)
    }
    Write-Host ""
    Write-Host "[DryRun] no files will be written"
    Write-Host "[DryRun] would sync    $source -> $staging"
    Write-Host "[DryRun] would install $staging -> $target"
    Write-Host "[DryRun] would set     dependencies.$packageName = $depValue"
    if ($hasLegacyDep) { Write-Host "[DryRun] would remove  dependencies.$legacyName" }
    Write-Host "[DryRun] would add     dsh.profile.bundles += $packageName"
    if ($hasLegacyBundle) { Write-Host "[DryRun] would remove  dsh.profile.bundles -= $legacyName" }
    if ($ExtraSectionPath -ne '') {
        Write-Host "[DryRun] would set     extraSections: $sectionName -> $ExtraSectionPath"
    }
    exit 0
}

function Invoke-Robocopy([string]$from, [string]$to) {
    robocopy $from $to /MIR /XD node_modules .git .superpowers /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed with exit code $LASTEXITCODE ($from -> $to)" }
    $global:LASTEXITCODE = 0
}

Write-Host "[1/5] staging  $source -> $staging"
New-Item -ItemType Directory -Force -Path $staging | Out-Null
Invoke-Robocopy $source $staging

Write-Host "[2/5] install  $staging -> $target"
if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
New-Item -ItemType Directory -Force -Path $target | Out-Null
Invoke-Robocopy $staging $target

# Machine-specific section paths stay OUT of the repo: the shipped
# cordis.patch.yml has `extraSections: []`, and this step fills in the local
# copy inside node_modules when the caller supplies one.
if ($ExtraSectionPath -ne '') {
    $patchPath = Join-Path $target 'cordis.patch.yml'
    if (Test-Path -LiteralPath $patchPath) {
        $name = $ExtraSectionName
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [IO.Path]::GetFileNameWithoutExtension($ExtraSectionPath) }
        $block = "        extraSections:`n          - name: '" + $name + "'`n            path: '" + $ExtraSectionPath + "'"
        $text = [IO.File]::ReadAllText($patchPath, $utf8)
        if ($text -match '(?m)^\s*extraSections:\s*\[\]\s*$') {
            $text = [regex]::Replace($text, '(?m)^\s*extraSections:\s*\[\]\s*$', $block)
            [IO.File]::WriteAllText($patchPath, $text, $utf8)
            Write-Host "  extraSections: $name -> $ExtraSectionPath"
        }
    }
}

Write-Host "[3/5] register in $pkgPath"
Copy-Item -LiteralPath $pkgPath -Destination "$pkgPath.bak-onenote-plugin" -Force
$pkg = [IO.File]::ReadAllText($pkgPath, $utf8) | ConvertFrom-Json

if ($null -eq $pkg.PSObject.Properties['dependencies']) {
    $pkg | Add-Member -NotePropertyName dependencies -NotePropertyValue ([pscustomobject]@{})
}
if ($pkg.dependencies.PSObject.Properties.Name -contains $legacyName) {
    $pkg.dependencies.PSObject.Properties.Remove($legacyName)
    Write-Host "  removed legacy dependency: $legacyName"
}
if ($pkg.dependencies.PSObject.Properties.Name -contains $packageName) {
    $pkg.dependencies.$packageName = $depValue
} else {
    $pkg.dependencies | Add-Member -NotePropertyName $packageName -NotePropertyValue $depValue
}

if ($null -eq $pkg.PSObject.Properties['dsh'] -or $null -eq $pkg.dsh.PSObject.Properties['profile']) {
    throw "unexpected profile package.json: missing dsh.profile"
}
$bundles = @($pkg.dsh.profile.bundles) | Where-Object { $_ -ne $legacyName }
if ($bundles -notcontains $packageName) { $bundles = @($bundles + $packageName) }
$pkg.dsh.profile.bundles = @($bundles)
if ($hasLegacyBundle) { Write-Host "  removed legacy bundle: $legacyName" }

$json = $pkg | ConvertTo-Json -Depth 12
[IO.File]::WriteAllText($pkgPath, $json, $utf8)

Write-Host "[4/5] pnpm install (keeps pnpm-lock.yaml in sync)"
$cli = $DshCli
if ($cli -eq '') {
    $command = Get-Command dsh -ErrorAction SilentlyContinue
    if ($null -ne $command) { $cli = $command.Source }
}
if ($cli -ne '' -and (Test-Path -LiteralPath $cli)) {
    $spec = $depValue
    $arguments = @('plugin', '--profile', $ProfileName, 'add')
    if ($Registry -ne '') { $arguments += @('--registry', $Registry) }
    $arguments += $spec
    & $cli @arguments
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "pnpm add failed (exit $LASTEXITCODE); the manual copy above still works, but re-run the pnpm add later to repair pnpm-lock.yaml."
    }
} else {
    Write-Warning "dsh CLI not found; skipping pnpm add. Run this once to sync pnpm-lock.yaml:"
    Write-Warning "  dsh plugin --profile $ProfileName add --registry $Registry `"$depValue`""
}

Write-Host "[5/5] done"
Write-Host ""
Write-Host "  dependency : $depValue"
Write-Host "  bundles    : $((@($pkg.dsh.profile.bundles)) -join ', ')"
Write-Host ""
Write-Host "Restart the DSH desktop app; the onenote_* tools appear after the restart (bundles are composed at startup)."
Write-Host "Uninstall: remove node_modules\onenote-plugin, the package.json dependency and bundles entry, then restart."
