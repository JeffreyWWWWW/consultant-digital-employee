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

function Copy-RootFile {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$DestinationRoot,
        [string]$BackupRoot
    )

    $source = Join-Path $SourceRoot $Name
    $destination = Join-Path $DestinationRoot $Name
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Missing source file: $source"
    }

    if ($BackupRoot -and (Test-Path -LiteralPath $destination)) {
        Copy-Item -LiteralPath $destination -Destination (Join-Path $BackupRoot $Name) -Force
    }
    Copy-Item -LiteralPath $source -Destination $destination -Force
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

    foreach ($name in @("AGENTS.md", "IDENTITY.md", "SOUL.md", "USER.md")) {
        $existing = Join-Path $OpenClawWorkspace $name
        if (Test-Path -LiteralPath $existing) {
            Copy-Item -LiteralPath $existing -Destination (Join-Path $backupRoot $name) -Force
        }
    }
    foreach ($dir in @("skills", "plugins")) {
        $existing = Join-Path $OpenClawWorkspace $dir
        if (Test-Path -LiteralPath $existing) {
            Copy-Item -LiteralPath $existing -Destination (Join-Path $backupRoot $dir) -Recurse -Force
        }
    }
}

foreach ($name in @("AGENTS.md", "IDENTITY.md", "SOUL.md", "USER.md")) {
    Copy-RootFile -Name $name -SourceRoot $configRoot -DestinationRoot $OpenClawWorkspace -BackupRoot $backupRoot
}

Invoke-RobocopyMirror -Source (Join-Path $configRoot "skills") -Destination (Join-Path $OpenClawWorkspace "skills")
Invoke-RobocopyMirror -Source (Join-Path $configRoot "plugins") -Destination (Join-Path $OpenClawWorkspace "plugins")

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
