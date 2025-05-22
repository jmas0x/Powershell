<#
.SYNOPSIS
    WSUS Update Manager
.DESCRIPTION
    Version 1.0
.Simulation mode (default)
    The script will only simulate the decline of updates without making any changes.
    
    .\wsus-arch-manager.ps1

.To process ARM only
    The script will decline all updates related to ARM architecture.
    .\wsus-arch-manager.ps1 -Architecture ARM -TrialRun $false

.For legacy systems (x86/32-bit)
    The script will decline all updates related to x86 architecture.
    .\wsus-arch-manager.ps1 -Architecture x86
#>

Param(
    [string]$WsusServer = $env:COMPUTERNAME,
    [bool]$UseSSL = $false,
    [int]$PortNumber = 8530,
    [bool]$TrialRun = $true,
    [ValidateSet('ARM','x86','Both')]
    [string]$Architecture = 'Both'
)

# Initialization
$global:startTime = Get-Date
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"

# Admin Check permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`nERROR: Execution requires administrative privileges!`n" -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit 1
}

# Just a Header
Write-Host "`n=== WSUS Architecture Manager v4.1 ===" -ForegroundColor Cyan
Write-Host "Target: $Architecture" -ForegroundColor Yellow
Write-Host "Server: $WsusServer" -ForegroundColor Yellow
Write-Host "Mode: $(if ($TrialRun) {'SIMULATION'} else {'EXECUTION'})" -ForegroundColor $(if ($TrialRun) {'Yellow'} else {'Red'})
Write-Host "Start: $($global:startTime.ToString('dd/MM/yyyy HH:mm:ss'))`n" -ForegroundColor Gray

# WSUS Connection
try {
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer, $UseSSL, $PortNumber)
    Write-Host "[STATUS] WSUS connection established successfully" -ForegroundColor Green
}
catch {
    Write-Host "`n[FAILURE] WSUS connection error:`n$($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "[DETAIL] $($_.Exception.InnerException.Message)" -ForegroundColor DarkYellow
    }
    exit 1
}

# Search Pattern
$pattern = switch ($Architecture) {
    'ARM'   { '(?i)\b(arm|aarch64)\b' }
    'x86'   { '(?i)\b(x86|32-bit|i386|win32)\b' }
    'Both'  { '(?i)\b(arm|aarch64|x86|32-bit|i386|win32)\b' }
}

# Main Processing
try {
    Write-Host "`n[STATUS] Searching for updates ($Architecture)..." -ForegroundColor Cyan
    
    $updates = $wsus.GetUpdates() | Where-Object {
        -not $_.IsDeclined -and 
        ($_.Title -match $pattern -or 
         $_.Description -match $pattern -or
         ($_.ProductTitles -match $pattern))
    }

    if (-not $updates) {
        Write-Host "[INFO] No updates for specified architectures were found.`n" -ForegroundColor Green
        exit 0
    }

    # Detailed Report
    $report = $updates | Select-Object @(
        @{N='Title';E={$_.Title}},
        @{N='Architecture';E={
            switch -regex ($_.Title + $_.Description) {
                '\b(ARM|aarch64)\b'    { 'ARM'; break }
                '\b(x86|32-bit|i386|win32)\b' { 'x86'; break }
                default      { 'Undetermined' }
            }
        }},
        @{N='Classification';E={$_.UpdateClassificationTitle}},
        @{N='KB';E={($_.KnowledgebaseArticles -join ', ').Trim(', ')}},
        @{N='Products';E={($_.ProductTitles -join ', ').Substring(0, [Math]::Min(30, ($_.ProductTitles -join ', ').Length)) + '...'}}
    )

    Write-Host "`n[RESULT] Distribution by Architecture:" -ForegroundColor Cyan
    $report | Group-Object Architecture | Sort-Object Count -Descending | Format-Table @(
        @{N='Architecture';E={$_.Name}},
        @{N='Count';E={$_.Count}},
        @{N='Example';E={$_.Group[0].Title.Substring(0, [Math]::Min(40, $_.Group[0].Title.Length)) + '...'}}
    ) -AutoSize

    # Decline Action
    if (-not $TrialRun) {
        Write-Host "`n[STATUS] Starting decline process..." -ForegroundColor Cyan
        $results = @{
            Success = 0
            Failures = 0
            Ignored = 0
        }

        $updates | ForEach-Object {
            try {
                $_.Decline()
                $results.Success++
                Write-Progress -Activity "Processing" -Status "$($results.Success)/$($updates.Count)" -PercentComplete (($results.Success/$updates.Count)*100)
            }
            catch {
                $results.Failures++
                Write-Host "  [WARNING] Failed to decline: $($_.Title)" -ForegroundColor DarkYellow
            }
        }

        Write-Host "`n[FINAL SUMMARY]" -ForegroundColor Cyan
        Write-Host "Successes: $($results.Success)" -ForegroundColor Green
        Write-Host "Failures: $($results.Failures)" -ForegroundColor $(if ($results.Failures -gt 0) {'Red'} else {'Gray'})
    }
    else {
        Write-Host "`n[INFO] Simulation mode - no actual changes were made`n" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n[CRITICAL ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    $endTime = Get-Date
    $duration = New-TimeSpan -Start $global:startTime -End $endTime
    Write-Host "`nTotal time: $($duration.ToString('hh\:mm\:ss\.fff'))" -ForegroundColor Cyan
    Write-Host "Completed at: $($endTime.ToString('dd/MM/yyyy HH:mm:ss'))`n" -ForegroundColor Gray
}