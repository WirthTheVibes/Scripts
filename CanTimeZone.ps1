$csv = Import-Csv <INPUT FILE HERE>

$total = $csv.Count
$count = 0
$percent = 0
$success = 0
$failed  = [System.Collections.Generic.List[object]]::new()
$remaining = "CALCULATING..."
$sw        = [System.Diagnostics.Stopwatch]::StartNew()

$migrationMap = @{
    'British Columbia Standard Time' = 'Pacific Standard Time'
    'Alberta Standard Time'          = 'Mountain Standard Time'
}

foreach ($item in $csv) {
    $count++
    $fromTz = $migrationMap[$item.timezone]
    Write-Host "[$count/$total - $percent%] Changing $($item.displayName) to $($item.timezone) - $remaining remaining." -ForegroundColor White -BackgroundColor Blue
    Write-Host
    try {
        $mailbox = Get-MailboxRegionalConfiguration -Identity $item.userPrincipalName
        Write-Host "Trying setting regional configuration..."
        if($mailbox.TimeZone -ne $item.timezone){
            Write-Host "Mailbox needs to be set to new time zone: $($item.timezone)." -ForegroundColor Yellow
            Set-MailboxRegionalConfiguration -Identity $item.userPrincipalName -TimeZone $item.timezone -ErrorAction Stop
            Set-MailboxCalendarConfiguration -Identity $item.userPrincipalName -WorkingHoursTimeZone $item.timezone
            Write-Host "Successfully set!" -ForegroundColor Green
        }
        else{
            Write-Host "No action required!" -ForegroundColor Green
        }
        Write-Host "Trying migrating calendar events..."
        if ($fromTz) {
            Invoke-CalendarEventMigration $item.userPrincipalName `
                -FromWindowsTimeZoneId $fromTz `
                -ToWindowsTimeZoneId   $item.timezone `
                -QueryStartDate        2026-11-01 `
                -QueryWindowInDays     365 `
                -Confirm:$false `
                -ErrorAction Stop
            Write-Host "Successfully set!" -ForegroundColor Green
        }
        $success++
    }
    catch {
        Write-Warning "Failed: $($item.userPrincipalName)"
        $failed.Add([PSCustomObject]@{
            DisplayName       = $item.displayName
            UserPrincipalName = $item.userPrincipalName
            TimeZone          = $item.timezone
            Error             = $_.Exception.Message
        })
    }
    $percent   = [math]::Round(($count / $total) * 100, 1)
    $avgSec    = $sw.Elapsed.TotalSeconds / $count
    $remaining = [TimeSpan]::FromSeconds($avgSec * ($total - $count)).ToString('hh\:mm\:ss')
}

Write-Host "`nDone. Success: $success  Failed: $($failed.Count)" -ForegroundColor Cyan

if ($failed.Count -gt 0) {
    Write-Host "`nFailed users:" -ForegroundColor Red
    $failed | Format-Table DisplayName, UserPrincipalName, TimeZone, Error -AutoSize -Wrap

    $failed | Export-Csv .\failedUsers_$(Get-Date -Format yyyy-MM-dd_HHmm).csv -NoTypeInformation
    Write-Host "Saved failed list to failedUsers CSV in the current folder."
}
