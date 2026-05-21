param(
    [string]$CommitMessage = "Update ppt-master OpenClaw skill"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$skillPath = Join-Path $repoRoot "openclaw-config\skills\ppt-master"

if (-not (Test-Path -LiteralPath $skillPath)) {
    throw "Missing submodule path: $skillPath"
}

Write-Host "Updating ppt-master skill submodule..."
git -C $skillPath pull origin openclaw-skill

Write-Host "Recording submodule pointer in main repo..."
git -C $repoRoot add openclaw-config\skills\ppt-master

$staged = git -C $repoRoot diff --cached --name-only -- openclaw-config/skills/ppt-master
if (-not $staged) {
    Write-Host "No submodule pointer change to commit."
    exit 0
}

git -C $repoRoot commit -m $CommitMessage -- openclaw-config\skills\ppt-master
Write-Host "Committed: $CommitMessage"
