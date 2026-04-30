param(
  [string]$Message = "Auto sync",
  [string]$Branch = "main",
  [int]$DebounceSeconds = 5,
  [switch]$Once
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$git = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Git.MinGit_Microsoft.Winget.Source_8wekyb3d8bbwe\cmd\git.exe"

if (!(Test-Path $git)) {
  $gitCommand = Get-Command git -ErrorAction SilentlyContinue
  if ($gitCommand) {
    $git = $gitCommand.Source
  } else {
    throw "Git was not found. Install Git, then run this script again."
  }
}

function Invoke-Git {
  & $git -C $root @args
}

function Sync-Changes {
  $changes = Invoke-Git status --porcelain
  if (!$changes) {
    return
  }

  Invoke-Git add -A

  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $commitMessage = "${Message}: $timestamp"
  Invoke-Git commit -m $commitMessage

  $remote = Invoke-Git remote
  if ($remote -contains "origin") {
    Invoke-Git push -u origin $Branch
  } else {
    Write-Host "No origin remote configured. Commit created locally: $commitMessage"
  }
}

Sync-Changes

if ($Once) {
  exit 0
}

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $root
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

$ignoredParts = @("\.git\", "\node_modules\", "\dist\", "\.output\", "\.tanstack\", "\.wrangler\")
$lastRun = Get-Date "2000-01-01"
$pending = $false

$action = {
  $path = $Event.SourceEventArgs.FullPath
  foreach ($part in $ignoredParts) {
    if ($path.Contains($part)) {
      return
    }
  }

  $script:pending = $true
}

$subscriptions = @(
  Register-ObjectEvent $watcher Changed -Action $action,
  Register-ObjectEvent $watcher Created -Action $action,
  Register-ObjectEvent $watcher Deleted -Action $action,
  Register-ObjectEvent $watcher Renamed -Action $action
)

Write-Host "Watching $root for changes. Press Ctrl+C to stop."

try {
  while ($true) {
    Start-Sleep -Seconds 1

    if (!$pending) {
      continue
    }

    $elapsed = ((Get-Date) - $lastRun).TotalSeconds
    if ($elapsed -lt $DebounceSeconds) {
      continue
    }

    $pending = $false
    $lastRun = Get-Date
    Sync-Changes
  }
} finally {
  foreach ($subscription in $subscriptions) {
    Unregister-Event -SubscriptionId $subscription.Id -ErrorAction SilentlyContinue
  }
  $watcher.Dispose()
}
