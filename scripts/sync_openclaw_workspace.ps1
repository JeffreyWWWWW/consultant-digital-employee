param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OpenClawWorkspace = (Join-Path $env:USERPROFILE ".openclaw\workspace"),
    [switch]$NoBackup,
    [switch]$RestartGateway
)

$ErrorActionPreference = "Stop"

function Invoke-RobocopyMirror {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Missing source directory: $Source"
    }

    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    robocopy $Source $Destination /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -gt 7) {
        throw "robocopy failed ($LASTEXITCODE): $Source -> $Destination"
    }
    $global:LASTEXITCODE = 0
}

function Restart-OpenClawGateway {
    $gatewayCmd = Join-Path $env:USERPROFILE ".openclaw\gateway.cmd"
    if (-not (Test-Path -LiteralPath $gatewayCmd)) {
        throw "Missing gateway command: $gatewayCmd"
    }

    Get-CimInstance Win32_Process |
        Where-Object { $_.Name -eq "node.exe" -and $_.CommandLine -like "*openclaw*" -and $_.CommandLine -like "*gateway*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }

    Start-Process -FilePath $gatewayCmd -WindowStyle Hidden
}

$configRoot = Join-Path $RepoRoot "openclaw-config"
if (-not (Test-Path -LiteralPath $configRoot)) {
    throw "Missing openclaw-config directory: $configRoot"
}

New-Item -ItemType Directory -Force -Path $OpenClawWorkspace | Out-Null

$backupRoot = $null
if (-not $NoBackup) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupRoot = Join-Path $env:USERPROFILE ".openclaw\backups\workspace-sync-$stamp"
    New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

    robocopy $OpenClawWorkspace $backupRoot /E /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -gt 7) {
        throw "backup failed ($LASTEXITCODE): $OpenClawWorkspace -> $backupRoot"
    }
    $global:LASTEXITCODE = 0
}

Invoke-RobocopyMirror -Source $configRoot -Destination $OpenClawWorkspace

if ($RestartGateway) {
    Restart-OpenClawGateway
}

Write-Host "OpenClaw workspace synchronized."
Write-Host "Source:      $configRoot"
Write-Host "Destination: $OpenClawWorkspace"
if ($backupRoot) {
    Write-Host "Backup:      $backupRoot"
}
if ($RestartGateway) {
    Write-Host "Gateway restart requested."
}
