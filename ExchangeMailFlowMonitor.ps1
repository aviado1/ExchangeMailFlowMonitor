<#
.SYNOPSIS
    Real-time Exchange Server 2016 Mail Flow Monitor.

.DESCRIPTION
    This PowerShell script monitors real-time mail flow and queue status 
    on an Exchange 2016 server. It displays the latest inbound and outbound
    emails as well as current queue information, refreshing at a configurable interval.

    Useful for system administrators who want a live snapshot of email activity 
    and mail queue status for troubleshooting or routine monitoring.

.NOTES
    Author: YourNameHere
    Version: 1.0
    GitHub: https://github.com/YourGitHubUsername

    Replace $server with your actual Exchange server name.
#>

# ======================
# CONFIGURATION SECTION
# ======================

$server = "EXCHANGE01"         # Replace with your Exchange server name
$refreshInterval = 10          # Interval in seconds between refreshes
$lookBackMinutes = 5           # Look back this many minutes for tracking logs

# Load Exchange snap-in if not already loaded
Try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
} Catch {
    Write-Host "Exchange snap-in already loaded or failed to load." -ForegroundColor Yellow
}

# ======================
# MONITORING LOOP
# ======================

while ($true) {
    Clear-Host
    $now = Get-Date
    $startTime = $now.AddMinutes(-$lookBackMinutes)
    Write-Host "===== Exchange Mail Traffic + Queue Monitor on $server (Updated: $now) =====" -ForegroundColor Cyan

    # === INBOUND MAIL TRACKING ===
    Write-Host "`n--- LAST 10 INBOUND MAILS ---" -ForegroundColor Green
    $inbound = Get-MessageTrackingLog -Server $server -Start $startTime -End $now `
        -EventId "RECEIVE" -ResultSize 100 |
        Where-Object { $_.Source -ne "STORE" } |
        Sort-Object Timestamp -Descending |
        Select-Object Timestamp, Sender, @{Name="Recipients"; Expression={$_.Recipients -join ","}}, MessageSubject -First 10

    if ($inbound) {
        $inbound | Format-Table -AutoSize
    } else {
        Write-Host "No recent inbound mail." -ForegroundColor Yellow
    }

    # === OUTBOUND MAIL TRACKING ===
    Write-Host "`n--- LAST 10 OUTBOUND MAILS ---" -ForegroundColor Yellow
    $outbound = Get-MessageTrackingLog -Server $server -Start $startTime -End $now `
        -EventId "SEND" -ResultSize 100 |
        Sort-Object Timestamp -Descending |
        Select-Object Timestamp, Sender, @{Name="Recipients"; Expression={$_.Recipients -join ","}}, MessageSubject -First 10

    if ($outbound) {
        $outbound | Format-Table -AutoSize
    } else {
        Write-Host "No recent outbound mail." -ForegroundColor Yellow
    }

    # === MAIL QUEUES ===
    Write-Host "`n--- CURRENT MAIL QUEUES ---" -ForegroundColor Magenta
    $queues = Get-Queue

    if ($queues.Count -eq 0) {
        Write-Host "No queues found." -ForegroundColor Yellow
    } else {
        $queues | Select Identity, Status, MessageCount, DeliveryType, NextHopDomain | Format-Table -AutoSize
    }

    # === MESSAGES IN QUEUE ===
    Write-Host "`n--- MESSAGES IN QUEUES ---" -ForegroundColor Blue
    $queuedMsgs = $queues | ForEach-Object {
        Get-Message -Queue $_.Identity -ErrorAction SilentlyContinue
    }

    if ($queuedMsgs.Count -eq 0) {
        Write-Host "No messages in any queue." -ForegroundColor Yellow
    } else {
        $queuedMsgs | Sort-Object DateReceived |
        Select Queue, FromAddress, Subject, Status, DateReceived |
        Format-Table -AutoSize
    }

    Start-Sleep -Seconds $refreshInterval
}
