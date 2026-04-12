[CmdletBinding()]
param(
    [string]$WorkbookPath = '',
    [string]$ReleaseDirectory = '',
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$projectRoot = Split-Path -Parent $scriptRoot
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $projectRoot 'AccountingIncOut.xlsm'
}
if ([string]::IsNullOrWhiteSpace($ReleaseDirectory)) {
    $ReleaseDirectory = Join-Path $projectRoot 'filesarchive\releases'
}

$resolvedWorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)
$resolvedReleaseDirectory = [System.IO.Path]::GetFullPath($ReleaseDirectory)

if (-not (Test-Path -LiteralPath $resolvedWorkbookPath)) {
    throw "Workbook not found: $resolvedWorkbookPath"
}

if (-not (Test-Path -LiteralPath $resolvedReleaseDirectory)) {
    New-Item -ItemType Directory -Path $resolvedReleaseDirectory -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedWorkbookPath)
$extension = [System.IO.Path]::GetExtension($resolvedWorkbookPath)
$releaseFileName = '{0}_{1}{2}' -f $baseName, $timestamp, $extension
$releasePath = Join-Path $resolvedReleaseDirectory $releaseFileName

if ($DryRun) {
    Write-Output "DRY_RUN: $releasePath"
    exit 0
}

Copy-Item -LiteralPath $resolvedWorkbookPath -Destination $releasePath -Force
Write-Output "RELEASE_CREATED: $releasePath"
