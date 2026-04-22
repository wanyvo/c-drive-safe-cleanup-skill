param(
    [string]$LogOutputPath = "",
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Format-Size {
    param([Int64]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    return ("{0:N2} GB" -f ($Bytes / 1GB))
}

function Get-DirectorySizeBytes {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return [Int64]0 }
    try {
        $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) { return [Int64]0 }
        return [Int64]$sum
    } catch {
        return [Int64]0
    }
}

function Resolve-ExistingPaths {
    param([string[]]$Patterns)
    $resolved = New-Object System.Collections.Generic.List[string]
    foreach ($pattern in $Patterns) {
        $expanded = [Environment]::ExpandEnvironmentVariables($pattern)
        if ($expanded.Contains("*")) {
            try {
                $items = Get-ChildItem -Path $expanded -Force -Directory -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    if ($item.FullName -like "C:\*") {
                        $resolved.Add($item.FullName)
                    }
                }
            } catch { }
        } else {
            if ((Test-Path -LiteralPath $expanded) -and ($expanded -like "C:\*")) {
                $resolved.Add($expanded)
            }
        }
    }
    return $resolved | Select-Object -Unique
}

function Get-AppDataResidualPaths {
    $roots = @(
        [Environment]::ExpandEnvironmentVariables("%APPDATA%"),
        [Environment]::ExpandEnvironmentVariables("%LOCALAPPDATA%")
    )
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        try {
            $dirs = Get-ChildItem -LiteralPath $root -Recurse -Directory -Force -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.FullName -like "C:\*" -and
                    $_.Name -match "(?i)(cache|tmp|temp|log|logs|backup)" -and
                    $_.FullName -notmatch "(?i)\\Microsoft\\|\\Windows\\|\\Packages\\"
                }
            foreach ($d in $dirs) {
                $result.Add($d.FullName)
            }
        } catch { }
    }
    return $result | Select-Object -Unique
}

function Clear-DirectoryContent {
    param([string]$Path)
    $errors = New-Object System.Collections.Generic.List[string]

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            DeletedBytes = [Int64]0
            Errors = @("Directory not found")
        }
    }

    $before = Get-DirectorySizeBytes -Path $Path
    try {
        $children = Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            try {
                Remove-Item -LiteralPath $child.FullName -Recurse -Force -ErrorAction Stop
            } catch {
                $errors.Add("Delete failed: $($child.FullName) | $($_.Exception.Message)")
            }
        }
    } catch {
        $errors.Add("Read directory failed: $Path | $($_.Exception.Message)")
    }

    $after = Get-DirectorySizeBytes -Path $Path
    return @{
        DeletedBytes = [Math]::Max([Int64]0, $before - $after)
        Errors = $errors
    }
}

function Test-WpsInstalled {
    $wpsExeCandidates = @(
        "C:\Program Files\WPS Office\*\WPS.exe",
        "C:\Program Files (x86)\WPS Office\*\WPS.exe"
    )
    foreach ($candidate in $wpsExeCandidates) {
        try {
            if (Get-ChildItem -Path $candidate -File -ErrorAction SilentlyContinue | Select-Object -First 1) {
                return $true
            }
        } catch { }
    }

    $wpsUserCandidates = @(
        "%APPDATA%\kingsoft\office6",
        "%LOCALAPPDATA%\Kingsoft\WPS Office"
    )
    foreach ($candidate in $wpsUserCandidates) {
        $p = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $p) { return $true }
    }

    return $false
}

function Get-BigFileReminders {
    param([int64]$ThresholdBytes = 500MB, [int]$TopN = 10)
    $targets = @(
        [Environment]::ExpandEnvironmentVariables("%USERPROFILE%\Downloads"),
        [Environment]::ExpandEnvironmentVariables("%USERPROFILE%\Desktop")
    )
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target)) { continue }
        try {
            $files = Get-ChildItem -LiteralPath $target -Recurse -File -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Length -ge $ThresholdBytes } |
                Sort-Object Length -Descending
            foreach ($f in $files) {
                $all.Add([PSCustomObject]@{
                    Path = $f.FullName
                    Size = [Int64]$f.Length
                })
            }
        } catch { }
    }
    return $all | Sort-Object Size -Descending | Select-Object -First $TopN
}

$wpsInstalled = Test-WpsInstalled

$entries = @(
    [PSCustomObject]@{
        Name = "User temp"
        Patterns = @("%TEMP%", "%LOCALAPPDATA%\Temp")
        Purpose = "Temporary working files for apps."
        WhyJunk = "Installers, updates and app crashes leave residue."
        Impact = "Safe to clean, first launch can be slightly slower."
    },
    [PSCustomObject]@{
        Name = "Windows temp"
        Patterns = @("%WINDIR%\Temp")
        Purpose = "Temporary files for OS and installers."
        WhyJunk = "Patch/install jobs leave stale files."
        Impact = "Mostly safe; locked files are skipped."
    },
    [PSCustomObject]@{
        Name = "Prefetch"
        Patterns = @("%WINDIR%\Prefetch")
        Purpose = "Startup optimization cache."
        WhyJunk = "Old records accumulate over time."
        Impact = "Rebuilt by Windows; short-term startup variance possible."
    },
    [PSCustomObject]@{
        Name = "INetCache"
        Patterns = @("%LOCALAPPDATA%\Microsoft\Windows\INetCache")
        Purpose = "Browser and network cache."
        WhyJunk = "Browsing and online usage continuously adds cache."
        Impact = "Safe; content will be downloaded again as needed."
    },
    [PSCustomObject]@{
        Name = "WER and CrashDumps"
        Patterns = @("%LOCALAPPDATA%\Microsoft\Windows\WER", "%LOCALAPPDATA%\CrashDumps")
        Purpose = "Error reports and crash dumps."
        WhyJunk = "Failures generate files that are rarely auto-removed."
        Impact = "Safe; fewer historic diagnostics remain."
    },
    [PSCustomObject]@{
        Name = "Store app cache"
        Patterns = @("%LOCALAPPDATA%\Packages\*\LocalCache")
        Purpose = "Cache for Microsoft Store apps."
        WhyJunk = "Apps generate temporary assets and local cache."
        Impact = "Safe; apps rebuild cache later."
    },
    [PSCustomObject]@{
        Name = "Delivery Optimization cache"
        Patterns = @("%PROGRAMDATA%\Microsoft\Windows\DeliveryOptimization\Cache")
        Purpose = "Windows update distribution cache."
        WhyJunk = "Update payload residue remains after updates."
        Impact = "Safe; does not remove installed updates."
    },
    [PSCustomObject]@{
        Name = "Graphics cache"
        Patterns = @("%LOCALAPPDATA%\D3DSCache", "%LOCALAPPDATA%\NVIDIA\DXCache", "%LOCALAPPDATA%\NVIDIA\GLCache")
        Purpose = "Shader and driver caches."
        WhyJunk = "Driver/app activity keeps generating data."
        Impact = "Safe; first graphics load may be slower."
    },
    [PSCustomObject]@{
        Name = "AppData residual cache"
        Patterns = @()
        Purpose = "Potential leftovers from uninstalled/unused apps."
        WhyJunk = "Uninstallers often leave cache/log folders."
        Impact = "Only cache-like names are included to reduce risk."
    }
)

if ($wpsInstalled) {
    $entries += [PSCustomObject]@{
        Name = "WPS cache/temp"
        Patterns = @(
            "%APPDATA%\kingsoft\office6\temp",
            "%LOCALAPPDATA%\Kingsoft\WPS Office\*cache*",
            "%LOCALAPPDATA%\Kingsoft\WPS Office\addons\pool\win-i386",
            "%LOCALAPPDATA%\Kingsoft\WPS Office\*temp*",
            "%LOCALAPPDATA%\Kingsoft\WPS Office\*backup*"
        )
        Purpose = "Temporary and cache files from WPS."
        WhyJunk = "Autosave/cache and update leftovers can grow."
        Impact = "Documents remain intact; temp history is removed."
    }
}

$scanResults = New-Object System.Collections.Generic.List[object]
foreach ($entry in $entries) {
    $paths = @(
        if ($entry.Name -eq "AppData residual cache") { Get-AppDataResidualPaths } else { Resolve-ExistingPaths -Patterns $entry.Patterns }
    )
    if (-not $paths -or $paths.Count -eq 0) { continue }

    $size = [Int64]0
    foreach ($p in $paths) { $size += Get-DirectorySizeBytes -Path $p }

    $scanResults.Add([PSCustomObject]@{
        Name = $entry.Name
        Paths = $paths
        Purpose = $entry.Purpose
        WhyJunk = $entry.WhyJunk
        Impact = $entry.Impact
        EstimatedBytes = $size
    })
}

$wpsScanItem = $scanResults | Where-Object { $_.Name -eq "WPS cache/temp" } | Select-Object -First 1
$wpsSummary = if ($null -ne $wpsScanItem) {
    "Detected cleanable WPS paths ($($wpsScanItem.Paths.Count) paths, estimated $(Format-Size $wpsScanItem.EstimatedBytes))"
} elseif ($wpsInstalled) {
    "Detected WPS installation traces, but no cleanable WPS cache paths found"
} else {
    "No WPS installation traces detected, WPS paths skipped"
}

if ($scanResults.Count -eq 0) {
    Write-Host "No cleanable whitelist paths found."
    exit 0
}

$actions = New-Object System.Collections.Generic.List[object]
foreach ($item in $scanResults) {
    Write-Host ""
    Write-Host "Group: $($item.Name)" -ForegroundColor Cyan
    Write-Host "Paths:" -ForegroundColor DarkCyan
    foreach ($p in $item.Paths) { Write-Host "  - $p" }
    Write-Host "Estimated free space: $(Format-Size $item.EstimatedBytes)"
    Write-Host "Purpose: $($item.Purpose)"
    Write-Host "Why junk grows: $($item.WhyJunk)"
    Write-Host "Impact: $($item.Impact)"

    $choice = Read-Host "Clean this group? (yes/no/skip)"
    $normalized = ""
    if ($null -ne $choice) { $normalized = $choice.Trim().ToLowerInvariant() }
    if ($normalized -in @("yes", "y")) {
        $actions.Add([PSCustomObject]@{ Item = $item; UserChoice = "yes" })
    } elseif ($normalized -in @("no", "n")) {
        $actions.Add([PSCustomObject]@{ Item = $item; UserChoice = "no" })
    } else {
        $actions.Add([PSCustomObject]@{ Item = $item; UserChoice = "skip" })
    }
}

$resultRows = New-Object System.Collections.Generic.List[object]
$totalReleased = [Int64]0
foreach ($action in $actions) {
    $item = $action.Item
    if ($action.UserChoice -ne "yes") {
        $resultRows.Add([PSCustomObject]@{
            Name = $item.Name
            Paths = ($item.Paths -join "<br>")
            UserChoice = $action.UserChoice
            Result = "not executed"
            ReleasedBytes = [Int64]0
            ErrorText = ""
        })
        continue
    }

    if ($DryRun) {
        $resultRows.Add([PSCustomObject]@{
            Name = $item.Name
            Paths = ($item.Paths -join "<br>")
            UserChoice = $action.UserChoice
            Result = "dry-run no delete"
            ReleasedBytes = [Int64]0
            ErrorText = ""
        })
        continue
    }

    $groupReleased = [Int64]0
    $errorMessages = New-Object System.Collections.Generic.List[string]
    foreach ($path in $item.Paths) {
        $clearResult = Clear-DirectoryContent -Path $path
        $groupReleased += [Int64]$clearResult.DeletedBytes
        foreach ($err in $clearResult.Errors) { $errorMessages.Add($err) }
    }
    $totalReleased += $groupReleased

    $status = if ($errorMessages.Count -eq 0) { "cleaned" } else { "partially failed" }
    $resultRows.Add([PSCustomObject]@{
        Name = $item.Name
        Paths = ($item.Paths -join "<br>")
        UserChoice = $action.UserChoice
        Result = $status
        ReleasedBytes = $groupReleased
        ErrorText = ($errorMessages -join " | ")
    })
}

$bigFiles = @(Get-BigFileReminders)
if ($bigFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "Reminder: large files in Downloads/Desktop (not auto-deleted)" -ForegroundColor Yellow
    foreach ($f in $bigFiles) { Write-Host ("- {0} ({1})" -f $f.Path, (Format-Size $f.Size)) }
}

$timestamp = Get-Date
$datePart = $timestamp.ToString("yyyy-MM-dd")
$timePart = $timestamp.ToString("yyyy-MM-dd HH:mm:ss")
if ([string]::IsNullOrWhiteSpace($LogOutputPath)) {
    $taskDir = Join-Path -Path (Get-Location) -ChildPath "tasks"
    if (-not (Test-Path -LiteralPath $taskDir)) { New-Item -ItemType Directory -Path $taskDir | Out-Null }
    $LogOutputPath = Join-Path -Path $taskDir -ChildPath ("evaluation-{0}-c-disk-cleanup.md" -f $datePart)
}

$policyText = if ($DryRun) {
    "whitelist only; per-group confirmation; dry-run no delete; no full-disk scan; no auto delete for Downloads/Desktop; no registry edits"
} else {
    "whitelist only; per-group confirmation; no full-disk scan; no auto delete for Downloads/Desktop; no registry edits"
}

$tableHeader = @(
    "| Group | Paths | User choice | Result | Freed space |",
    "|---|---|---|---|---|"
)
$tableRows = foreach ($row in $resultRows) {
    $safeResult = if ([string]::IsNullOrWhiteSpace($row.ErrorText)) { $row.Result } else { "{0} ({1})" -f $row.Result, $row.ErrorText }
    "| $($row.Name) | $($row.Paths) | $($row.UserChoice) | $safeResult | $(Format-Size $row.ReleasedBytes) |"
}

$bigFileSection = if ($bigFiles.Count -gt 0) {
    $lines = @("## Large file reminders (not deleted)", "")
    foreach ($f in $bigFiles) { $lines += "- $($f.Path) ($(Format-Size $f.Size))" }
    $lines -join "`n"
} else {
    "## Large file reminders (not deleted)`n`n- None above threshold."
}

$markdown = @"
# C Drive Cleanup Log

- Time: $timePart
- Drive: C:
- Policy: $policyText

## Details

$(($tableHeader + $tableRows) -join "`n")

## Summary

- Total released: $(Format-Size $totalReleased)
- WPS detection: $wpsSummary
- Run mode: $(if ($DryRun) { "DryRun (no deletion)" } else { "Actual cleanup" })

$bigFileSection
"@

Set-Content -LiteralPath $LogOutputPath -Value $markdown -Encoding UTF8

Write-Host ""
if ($DryRun) {
    Write-Host "Dry run completed. No files were deleted." -ForegroundColor Yellow
} else {
    Write-Host "Cleanup completed. Total released: $(Format-Size $totalReleased)" -ForegroundColor Green
}
Write-Host "Log file: $LogOutputPath" -ForegroundColor Green
