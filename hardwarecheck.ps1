<#
.SYNOPSIS
    Apollo Technology Hardware Diagnostics Utility v1.0
.DESCRIPTION
    Menu-driven utility to automatically scan, verify, and report on system hardware health.
    - INCLUDES: CPU, RAM, Disk Health, GPU, Network, USB Error Checking, Battery Health.
    - FEATURES: Anti-Sleep, Anti-Freeze, PDF Edge Reporting, Email Delivery.
#>

# --- 0. CONFIGURATION ---
$VerboseMode = $false         # Set to $true to log all script output to C:\temp\hwcheck
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$Version = "1.0"

# --- EMAIL SETTINGS ---
$EmailEnabled = $false       # Set to $true to enable email
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$FromAddress  = "reports@yourdomain.com"
$ToAddress    = "support@yourdomain.com"
$UseSSL       = $true

# --- 1. AUTO-ELEVATE TO ADMINISTRATOR ---
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (!($isAdmin)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    try {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
    catch {
        Write-Error "Failed to elevate. Please run as Administrator manually."
        Pause
        Exit
    }
}

# --- 1.5 VERBOSE LOGGING SETUP & WARNING ---
if ($VerboseMode) {
    $VerboseDir = "C:\temp\hwcheck"
    if (!(Test-Path $VerboseDir)) {
        New-Item -ItemType Directory -Path $VerboseDir -Force | Out-Null
    }
    
    $LogTimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $TranscriptPath = "$VerboseDir\hwcheck_logs_$LogTimeStamp.txt"
    
    Start-Transcript -Path $TranscriptPath -Force | Out-Null
    
    Clear-Host
    Write-Host "`n=================================================================================" -ForegroundColor Magenta
    Write-Host " [ WARNING: VERBOSE LOGGING IS ENABLED ]" -ForegroundColor Red
    Write-Host "=================================================================================" -ForegroundColor Magenta
    Write-Host " All console output and errors are currently being recorded."
    Write-Host " Log File Location: " -NoNewline; Write-Host $TranscriptPath -ForegroundColor Cyan
    Write-Host "`n---------------------------------------------------------------------------------" -ForegroundColor DarkGray
    $null = Read-Host " Press [ENTER] to acknowledge and continue"
}

# --- 2. PREVENT FREEZING & SLEEPING ---
$consoleFuncs = @"
using System;
using System.Runtime.InteropServices;
public class ConsoleUtils {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void DisableQuickEdit() {
        IntPtr hConsole = GetStdHandle(-10);
        uint mode;
        GetConsoleMode(hConsole, out mode);
        mode &= ~0x0040u;
        SetConsoleMode(hConsole, mode);
    }
}
"@
try { Add-Type -TypeDefinition $consoleFuncs -Language CSharp; [ConsoleUtils]::DisableQuickEdit() } catch { }

$sleepBlocker = @"
using System;
using System.Runtime.InteropServices;
public class SleepUtils {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@
try { Add-Type -TypeDefinition $sleepBlocker -Language CSharp; $null = [SleepUtils]::SetThreadExecutionState(0x80000003) } catch { }

# --- 3. HELPER FUNCTIONS ---
function Show-Header {
    Clear-Host
    $Banner = @'
   __  _____    ____  ____ _       __ ___    ____  ______
  / / / /   |  / __ \/ __ \ |     / //   |  / __ \/ ____/
 / /_/ / /| | / /_/ / / / / | /| / // /| | / /_/ / __/   
/ __  / ___ |/ _, _/ /_/ /| |/ |/ // ___ |/ _, _/ /___   
/_/ /_/_/  |_/_/ |_/_____/ |__/|__/_/  |_/_/ |_/_____/   
'@
    Write-Host $Banner -ForegroundColor Cyan
    Write-Host "`n   HARDWARE DIAGNOSTICS TOOL v$Version" -ForegroundColor White
    Write-Host "=================================================================================" -ForegroundColor DarkGray
    Write-Host "        [NOTICE] Running in Elevated Permissions" -ForegroundColor Red 
    Write-Host "      [POWER] Sleep Mode & Screen Timeout Blocked." -ForegroundColor DarkGray
}

# --- 4. MAIN MENU & INPUTS ---
Show-Header
Write-Host "`n[ CONFIGURATION ]" -ForegroundColor Yellow

$EngineerName = Read-Host "   > Enter Engineer Name"
$TicketNumber = Read-Host "   > Enter Ticket Number"
$CustomerName = Read-Host "   > Enter Customer Name"

$EmailCreds = $null
if ($EmailEnabled) {
    Write-Host "`n[ EMAIL CONFIGURATION ]" -ForegroundColor Cyan
    $EmailPass = Read-Host "   > Please enter the Password for $FromAddress" -AsSecureString
    $EmailCreds = New-Object System.Management.Automation.PSCredential ($FromAddress, $EmailPass)
}

# --- 5. HARDWARE CHECKS ---
$ReportItems = @()
$RunTimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportDir = "C:\HardwareReports"
if (!(Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }

Show-Header
Write-Host "`n[ EXECUTING HARDWARE DIAGNOSTICS ]" -ForegroundColor Yellow

$DiagnosticsList = @("System Info", "CPU Status", "Memory (RAM)", "Disk Drives", "Display / GPU", "Network Adapters", "USB Devices", "Battery (If Present)")
$TotalTasks = $DiagnosticsList.Count
$CurrentTask = 0

function Add-Result($Component, $Details, $Status) {
    $script:ReportItems += [PSCustomObject]@{ Component = $Component; Details = $Details; Status = $Status }
}

# 1. System Info
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[0])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking System Information..." -ForegroundColor Green
$SysInfo = Get-CimInstance Win32_ComputerSystem
$Bios = Get-CimInstance Win32_BIOS
Add-Result "System" "Model: $($SysInfo.Model) | BIOS: $($Bios.SMBIOSBIOSVersion)" "OK"
Start-Sleep -Milliseconds 500

# 2. CPU
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[1])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking CPU Status..." -ForegroundColor Green
$CPU = Get-CimInstance Win32_Processor
$CpuStatus = if ($CPU.Status -eq "OK") { "Pass" } else { "Warning" }
Add-Result "CPU" "Name: $($CPU.Name) | Cores: $($CPU.NumberOfCores) | Load: $($CPU.LoadPercentage)%" $CpuStatus
Start-Sleep -Milliseconds 500

# 3. RAM
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[2])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Physical Memory..." -ForegroundColor Green
$RAM = Get-CimInstance Win32_PhysicalMemory
$TotalRamGB = [math]::Round(($RAM | Measure-Object Capacity -Sum).Sum / 1GB, 2)
$MemStatus = if ($TotalRamGB -ge 4) { "Pass" } else { "Low Memory" }
Add-Result "RAM" "Total Installed: ${TotalRamGB}GB | Sticks: $($RAM.Count)" $MemStatus
Start-Sleep -Milliseconds 500

# 4. Disks
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[3])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Disk SMART/Health Status..." -ForegroundColor Green
$Disks = Get-CimInstance Win32_DiskDrive
foreach ($Disk in $Disks) {
    $SizeGB = [math]::Round($Disk.Size / 1GB, 2)
    $Status = if ($Disk.Status -eq "OK") { "Pass" } else { "FAIL" }
    Add-Result "Disk" "$($Disk.Model) (${SizeGB}GB)" $Status
}
Start-Sleep -Milliseconds 500

# 5. GPU
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[4])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Display Adapters..." -ForegroundColor Green
$GPUs = Get-CimInstance Win32_VideoController
foreach ($GPU in $GPUs) {
    $Status = if ($GPU.Status -eq "OK") { "Pass" } else { "Error" }
    Add-Result "GPU" "$($GPU.Name) | Driver: $($GPU.DriverVersion)" $Status
}
Start-Sleep -Milliseconds 500

# 6. Network
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[5])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Network Adapters..." -ForegroundColor Green
$Nets = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
foreach ($Net in $Nets) {
    $NetStatus = if ($Net.NetConnectionStatus -eq 2) { "Connected" } elseif ($Net.NetConnectionStatus -eq 7) { "Disconnected" } else { "Disabled/Other" }
    Add-Result "Network" "$($Net.Name)" $NetStatus
}
Start-Sleep -Milliseconds 500

# 7. USB Errors
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[6])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking USB Device Tree for Errors..." -ForegroundColor Green
$USBErrors = Get-CimInstance Win32_PnPEntity | Where-Object { $_.DeviceID -match "USB" -and $_.ConfigManagerErrorCode -ne 0 }
if ($USBErrors) {
    foreach ($err in $USBErrors) { Add-Result "USB" "$($err.Name) (Error Code: $($err.ConfigManagerErrorCode))" "FAIL" }
} else {
    Add-Result "USB" "No USB controller or device errors detected." "Pass"
}
Start-Sleep -Milliseconds 500

# 8. Battery
$CurrentTask++; Write-Progress -Activity "Scanning Hardware" -Status "Checking $($DiagnosticsList[7])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Battery..." -ForegroundColor Green
$Battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
if ($Battery) {
    Add-Result "Battery" "Status: $($Battery.Status) | Charge: $($Battery.EstimatedChargeRemaining)%" "Pass"
} else {
    Add-Result "Battery" "No battery detected (Desktop PC)." "N/A"
}

Write-Progress -Activity "Scanning Hardware" -Completed

# --- 6. REPORT GENERATION ---
Write-Host "`n[ REPORT GENERATION ]" -ForegroundColor Yellow
$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$TransferTableRows = ""

foreach ($Row in $ReportItems) {
    $StatusColor = switch -Regex ($Row.Status) {
        "Pass|OK|Connected" { "green" }
        "Warning|Low Memory|Disconnected" { "darkorange" }
        "FAIL|Error" { "red" }
        Default { "#555" }
    }
    $TransferTableRows += "<tr><td><strong>$($Row.Component)</strong></td><td>$($Row.Details)</td><td><strong style='color:$StatusColor'>$($Row.Status)</strong></td></tr>"
}

$HtmlFile = "$ReportDir\HW_Report_${TicketNumber}_$RunTimeStamp.html"
$PdfFile  = "$ReportDir\HW_Report_${TicketNumber}_$RunTimeStamp.pdf"

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', sans-serif; color: #333; padding: 20px; }
    .header { text-align: center; margin-bottom: 20px; }
    h1 { color: #0056b3; margin-bottom: 5px; }
    .meta { font-size: 0.9em; color: #666; text-align: center; margin-bottom: 30px; }
    .section { background: #f9f9f9; padding: 15px; border-left: 6px solid #0056b3; margin-bottom: 20px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
    th { text-align: left; background: #eee; padding: 8px; border-bottom: 1px solid #ddd; }
    td { padding: 8px; border-bottom: 1px solid #ddd; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" style="max-height:100px;">
    <h1>Hardware Diagnostics Report</h1>
    <p>Report generated by <strong>$EngineerName</strong> for ticket (<strong>$TicketNumber</strong>)</p>
    <div class="meta">
        <strong>Date:</strong> $CurrentDate | <strong>Customer:</strong> $CustomerName
    </div>
</div>
<h2>Workstation Information</h2>
<div class="section">
    <strong>Computer Name:</strong> $env:COMPUTERNAME <br>
    <strong>OS Version:</strong> $((Get-CimInstance Win32_OperatingSystem).Caption) <br>
    <strong>Engineer:</strong> $EngineerName
</div>
<h2>Diagnostics Results</h2>
<div class="section">
    <table><thead><tr><th>Component</th><th>Details</th><th>Status</th></tr></thead><tbody>$TransferTableRows</tbody></table>
</div>
<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $(Get-Date -Format yyyy) by Apollo Technology.</p>
</body>
</html>
"@

$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8

# Convert to PDF
$EdgeLoc1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$EdgeLoc2 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$EdgeExe = if (Test-Path $EdgeLoc1) { $EdgeLoc1 } elseif (Test-Path $EdgeLoc2) { $EdgeLoc2 } else { $null }

if ($EdgeExe) {
    Write-Host "   Generating PDF Report..." -ForegroundColor Cyan
    $EdgeUserData = "$ReportDir\EdgeTemp"
    if (-not (Test-Path $EdgeUserData)) { New-Item -Path $EdgeUserData -ItemType Directory -Force | Out-Null }
    try {
        Start-Process -FilePath $EdgeExe -ArgumentList "--headless", "--print-to-pdf=`"$PdfFile`"", "--no-pdf-header-footer", "--user-data-dir=`"$EdgeUserData`"", "`"$HtmlFile`"" -Wait
        if (Test-Path $PdfFile) {
            Write-Host "   Report saved to: $PdfFile" -ForegroundColor Green
            Remove-Item $HtmlFile -ErrorAction SilentlyContinue
            Remove-Item $EdgeUserData -Recurse -Force -ErrorAction SilentlyContinue
            Start-Process $PdfFile
        }
    } catch {
        Write-Warning "PDF Conversion failed. Report saved as HTML."
        $PdfFile = $HtmlFile
        Start-Process $HtmlFile
    }
} else {
    Write-Warning "Edge not found. Report saved as HTML."
    $PdfFile = $HtmlFile
    Start-Process $HtmlFile
}

# --- 7. EMAIL REPORT ---
if ($EmailEnabled -and $PdfFile -and (Test-Path $PdfFile)) {
    Write-Host "`n[ EMAIL REPORT ]" -ForegroundColor Yellow
    Write-Host "   Sending Email to $ToAddress..." -ForegroundColor Cyan
    try {
        Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Hardware Check: $env:COMPUTERNAME ($TicketNumber)" -Body "Attached is the hardware report for Ticket $TicketNumber ($CustomerName)." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PdfFile -ErrorAction Stop
        Write-Host "   > Email Sent Successfully!" -ForegroundColor Green
    } catch {
        Write-Error "   > Failed to send email. Error: $_"
    }
}

# --- ALLOW SLEEP AGAIN ---
try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Diagnostics finished. The report has been opened."

if ($VerboseMode) {
    Write-Host "Stopping Verbose Logging..." -ForegroundColor DarkGray
    Stop-Transcript | Out-Null
}

Pause