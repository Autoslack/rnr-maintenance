# =========================================================
# COMPANY ENTERPRISE MAINTENANCE DEPLOYMENT - v6 FINAL
# Includes Saturday Warning + 9PM Cleanup + 10PM Reboot
# =========================================================

Write-Host "Deploying Enterprise Maintenance Tasks..." -ForegroundColor Cyan

$scriptRoot = "C:\CompanyScripts"
New-Item -ItemType Directory -Path $scriptRoot -Force | Out-Null

# ---------------------------------------------------------
# COMMON TASK SETTINGS (Reliability Hardening)
# ---------------------------------------------------------

$settings = New-ScheduledTaskSettingsSet `
-AllowStartIfOnBatteries `
-DontStopIfGoingOnBatteries `
-StartWhenAvailable `
-RestartCount 3 `
-RestartInterval (New-TimeSpan -Minutes 5)

# ---------------------------------------------------------
# SATURDAY 4PM WARNING POPUP SCRIPT
# ---------------------------------------------------------

$warningScript = @'
msg * "NOTICE: All company computers will perform scheduled maintenance and reboot tonight at 10:00 PM. Please save your work before leaving and leave your computer powered ON."
'@

$warningScript | Set-Content "$scriptRoot\Saturday-Warning.ps1"

# ---------------------------------------------------------
# WEEKLY CLEANUP SCRIPT (Saturday 9PM)
# ---------------------------------------------------------

$cleanupScript = @'
$logFile = "C:\CompanyScripts\MaintenanceLog.txt"
Add-Content $logFile "$(Get-Date) - Weekly Cleanup Started"

# Clear Windows temp
Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

# Clear user temp folders
Get-ChildItem "C:\Users" -Directory | ForEach-Object {
    $tempPath = "$($_.FullName)\AppData\Local\Temp"
    Remove-Item "$tempPath\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# Clear root C:\Downloads (if exists)
if (Test-Path "C:\Downloads") {
    Remove-Item "C:\Downloads\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# Clear Windows Update cache
Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
Start-Service wuauserv -ErrorAction SilentlyContinue

# Clear Recycle Bin
Clear-RecycleBin -Force -ErrorAction SilentlyContinue

# Component store cleanup
DISM /Online /Cleanup-Image /StartComponentCleanup /Quiet

Add-Content $logFile "$(Get-Date) - Weekly Cleanup Completed"
'@

$cleanupScript | Set-Content "$scriptRoot\Weekly-Cleanup.ps1"

# ---------------------------------------------------------
# DAILY CHROME CACHE PURGE (6AM - Password Safe)
# ---------------------------------------------------------

$chromeScript = @'
$logFile = "C:\CompanyScripts\MaintenanceLog.txt"
Add-Content $logFile "$(Get-Date) - Daily Chrome Cache Purge Started"

Get-ChildItem "C:\Users" -Directory | ForEach-Object {
    $base = "$($_.FullName)\AppData\Local\Google\Chrome\User Data\Default"
    
    $paths = @(
        "$base\Cache",
        "$base\Code Cache",
        "$base\GPUCache",
        "$base\Media Cache"
    )

    foreach ($path in $paths) {
        Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Add-Content $logFile "$(Get-Date) - Daily Chrome Cache Purge Completed"
'@

$chromeScript | Set-Content "$scriptRoot\Daily-ChromePurge.ps1"

# ---------------------------------------------------------
# EVENT LOG CLEANUP (Monthly)
# ---------------------------------------------------------

$eventScript = @'
$logFile = "C:\CompanyScripts\MaintenanceLog.txt"
Add-Content $logFile "$(Get-Date) - Event Log Cleanup Started"
wevtutil el | ForEach-Object { wevtutil cl "$_" }
Add-Content $logFile "$(Get-Date) - Event Log Cleanup Completed"
'@

$eventScript | Set-Content "$scriptRoot\EventLog-Cleanup.ps1"

# ---------------------------------------------------------
# UPTIME LOGGER (Startup)
# ---------------------------------------------------------

$uptimeScript = @'
$logFile = "C:\CompanyScripts\UptimeLog.txt"
$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
Add-Content $logFile "$(Get-Date) - Boot detected. Last Boot: $boot"
'@

$uptimeScript | Set-Content "$scriptRoot\Startup-Uptime.ps1"

# ---------------------------------------------------------
# AUTO-START CHROME + OUTLOOK + GOOGLE CHAT
# ---------------------------------------------------------

$startupFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
$wshell = New-Object -ComObject WScript.Shell

# Chrome
$chromeExe = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (Test-Path $chromeExe) {
    $shortcut = $wshell.CreateShortcut("$startupFolder\Google Chrome.lnk")
    $shortcut.TargetPath = $chromeExe
    $shortcut.Save()
}

# Outlook
$outlookExe = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
if (Test-Path $outlookExe) {
    $shortcut = $wshell.CreateShortcut("$startupFolder\Microsoft Outlook.lnk")
    $shortcut.TargetPath = $outlookExe
    $shortcut.Save()
}

# Google Chat (if installed)
$chatPaths = @(
    "C:\Program Files\Google\Chat\GoogleChat.exe",
    "C:\Program Files (x86)\Google\Chat\GoogleChat.exe"
)

foreach ($path in $chatPaths) {
    if (Test-Path $path) {
        $shortcut = $wshell.CreateShortcut("$startupFolder\Google Chat.lnk")
        $shortcut.TargetPath = $path
        $shortcut.Save()
        break
    }
}

# ---------------------------------------------------------
# FUNCTION TO REGISTER TASK
# ---------------------------------------------------------

function Register-CompanyTask {
    param($Name,$Script,$Trigger)

    $action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -File `"$Script`""

    Register-ScheduledTask `
    -TaskName $Name `
    -Action $action `
    -Trigger $Trigger `
    -User "SYSTEM" `
    -RunLevel Highest `
    -Settings $settings `
    -Force
}

# ---------------------------------------------------------
# TASK SCHEDULING
# ---------------------------------------------------------

# Saturday 4PM - Maintenance Warning
Register-CompanyTask `
"Company Saturday Maintenance Warning" `
"$scriptRoot\Saturday-Warning.ps1" `
(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -At 4:00PM)

# Saturday 9PM - Cleanup
Register-CompanyTask `
"Company Weekly Cleanup" `
"$scriptRoot\Weekly-Cleanup.ps1" `
(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -At 9:00PM)

# Saturday 10PM - Reboot
$rebootAction = New-ScheduledTaskAction -Execute "shutdown.exe" `
-Argument "/r /f /t 300 /c `"Weekly Maintenance Reboot - Save Work Now`""

Register-ScheduledTask `
-TaskName "Company Weekly Reboot" `
-Action $rebootAction `
-Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -At 10:00PM) `
-User "SYSTEM" `
-RunLevel Highest `
-Settings $settings `
-Force

# Daily 6AM - Chrome Cache Purge
Register-CompanyTask `
"Company Daily Chrome Cache Purge" `
"$scriptRoot\Daily-ChromePurge.ps1" `
(New-ScheduledTaskTrigger -Daily -At 6:00AM)

# Sunday 1AM - Windows Update Scan
Register-CompanyTask `
"Company Weekly Update Scan" `
"$scriptRoot\Weekly-Cleanup.ps1" `
(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 1:00AM)

# Sunday 2AM - Defender Scan
Register-CompanyTask `
"Company Weekly Defender Scan" `
"$scriptRoot\Weekly-Cleanup.ps1" `
(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2:00AM)

# Sunday 3AM - Disk Optimize
Register-CompanyTask `
"Company Weekly Disk Optimize" `
"$scriptRoot\Weekly-Cleanup.ps1" `
(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 3:00AM)

# First Sunday 4AM - Event Log Cleanup
Register-CompanyTask `
"Company Monthly EventLog Cleanup" `
"$scriptRoot\EventLog-Cleanup.ps1" `
(New-ScheduledTaskTrigger -Monthly -DaysOfWeek Sunday -WeeksOfMonth First -At 4:00AM)

# Startup Uptime Logging
Register-CompanyTask `
"Company Startup Uptime Logger" `
"$scriptRoot\Startup-Uptime.ps1" `
(New-ScheduledTaskTrigger -AtStartup)

Write-Host "Enterprise Maintenance Deployment Complete." -ForegroundColor Green
