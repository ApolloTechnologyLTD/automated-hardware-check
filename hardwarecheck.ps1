<#
.SYNOPSIS
    Apollo Technology Ultimate Hardware & OS Diagnostics v3.4
.DESCRIPTION
    The most comprehensive native PowerShell hardware diagnostic tool.
    - INCLUDES: TPM, Secure Boot, Firewall, IP/Wi-Fi, Audio, Universal Device Errors, Logical Volumes.
    - PATCHED: Subexpression parser bug in System & Boot.
    - PATCHED: Date conversion error for CimInstance objects.
    - PATCHED: Divide-by-zero errors for ghost/unformatted partitions.
    - RESTORED: Verbose Transcript logging for background debugging.
    - TWEAKED: Removed bullet points to fix encoding artifacts in PDF generation.
    - TWEAKED: Updated output directory and footer text.
#>

# --- 0. CONFIGURATION ---
$VerboseMode  = $false        # Set to $true to log raw background processes to C:\temp\hwcheck
$LogoUrl      = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$Version      = "3.7"
$ReportDir    = "C:\temp\Apollo_Reports"

# --- EMAIL SETTINGS ---
$EmailEnabled = $false
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$FromAddress  = "reports@yourdomain.com"
$ToAddress    = "support@yourdomain.com"
$UseSSL       = $true

# --- 1. AUTO-ELEVATE TO ADMINISTRATOR ---
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "Requesting Administrator privileges for deep hardware access..." -ForegroundColor Yellow
    try { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit }
    catch { Write-Error "Failed to elevate. Please run as Administrator manually."; Pause; Exit }
}

# --- 1.5 VERBOSE LOGGING SETUP ---
if ($VerboseMode) {
    $VerboseDir = "C:\temp\hwcheck"
    if (!(Test-Path $VerboseDir)) { New-Item -ItemType Directory -Path $VerboseDir -Force | Out-Null }
    
    $LogTimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $TranscriptPath = "$VerboseDir\hwcheck_verbose_$LogTimeStamp.txt"
    
    Start-Transcript -Path $TranscriptPath -Force | Out-Null
    
    Clear-Host
    Write-Host "`n=================================================================================" -ForegroundColor Magenta
    Write-Host " [ WARNING: VERBOSE LOGGING IS ENABLED ]" -ForegroundColor Red
    Write-Host "=================================================================================" -ForegroundColor Magenta
    Write-Host " All console output and background errors are currently being recorded."
    Write-Host " Log File Location: " -NoNewline; Write-Host $TranscriptPath -ForegroundColor Cyan
    Write-Host "`n---------------------------------------------------------------------------------" -ForegroundColor DarkGray
    $null = Read-Host " Press [ENTER] to acknowledge and continue"
}

# --- 2. PREVENT FREEZING & SLEEPING ---
$consoleFuncs = @"
using System; using System.Runtime.InteropServices;
public class ConsoleUtils {
    [DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void DisableQuickEdit() {
        IntPtr hConsole = GetStdHandle(-10); uint mode;
        GetConsoleMode(hConsole, out mode); mode &= ~0x0040u; SetConsoleMode(hConsole, mode);
    }
}
"@
try { Add-Type -TypeDefinition $consoleFuncs -Language CSharp; [ConsoleUtils]::DisableQuickEdit() } catch { }

$sleepBlocker = @"
using System; using System.Runtime.InteropServices;
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
    __  ____  __  _____                 __       ____  _                 
   / / / / / / /_  __/ (_)___ ___  ____ _/ /____  / __ \(_)___ _____ _____ 
  / / / / / / / / /   / / __ `__ \/ __ `/ __/ _ \/ / / / / __ `/ __ `/ ___/
 / /_/ / /_/ / / /   / / / / / / / /_/ / /_/  __/ /_/ / / /_/ / /_/ (__  ) 
 \____/\____/ /_/   /_/_/ /_/ /_/\__,_/\__/\___/_____/_/\__,_/\__, /____/  
                                                             /____/        
'@
    Write-Host $Banner -ForegroundColor Cyan
    Write-Host "`n   ULTIMATE HARDWARE DIAGNOSTICS TOOL v$Version" -ForegroundColor White
    Write-Host "=================================================================================" -ForegroundColor DarkGray
    Write-Host "        [NOTICE] Running in Elevated Permissions" -ForegroundColor Red 
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
if (!(Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }

Show-Header
Write-Host "`n[ EXECUTING ULTIMATE HARDWARE DIAGNOSTICS ]" -ForegroundColor Yellow

$DiagnosticsList = @("System & Boot", "OS & Security", "Processor & Board", "Memory & Virtual", "Physical Storage", "Logical Volumes", "Graphics & Audio", "Network & Comms", "Hardware Health")
$TotalTasks = $DiagnosticsList.Count
$CurrentTask = 0

function Add-Result($Category, $Component, $Details, $Status) {
    $script:ReportItems += [PSCustomObject]@{ Category = $Category; Component = $Component; Details = $Details; Status = $Status }
}

# 1. System & Boot
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[0])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Firmware, Boot & TPM..." -ForegroundColor Green

$Bios = Get-CimInstance Win32_BIOS
$BiosDate = if ($Bios.ReleaseDate) { $Bios.ReleaseDate.ToString('yyyy-MM-dd') } else { "Unknown" }

try { $SecureBoot = Confirm-SecureBootUEFI -ErrorAction Stop; $SBStatus = if ($SecureBoot) { "Enabled" } else { "Disabled" } } catch { $SBStatus = "Unsupported/Legacy" }
try { $TPM = Get-Tpm -ErrorAction Stop; $TpmStatus = if ($TPM.TpmPresent) { "Ready (v$($TPM.TpmReady))" } else { "Not Present" } } catch { $TpmStatus = "Not Supported" }

$SBStatusColor = if ($SBStatus -eq "Enabled") {"Pass"} else {"Warning"}
$TpmStatusColor = if ($TpmStatus -match "Ready") {"Pass"} else {"Warning"}

Add-Result "System & Boot" "Firmware & BIOS" "<b>Version:</b> $($Bios.SMBIOSBIOSVersion)<br><b>Release Date:</b> $BiosDate<br><b>Secure Boot:</b> $SBStatus" $SBStatusColor
Add-Result "System & Boot" "Security Chip" "<b>TPM Status:</b> $TpmStatus" $TpmStatusColor

# 2. OS & Security Check
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[1])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking OS, Updates, and Security Policies..." -ForegroundColor Green
$OS = Get-CimInstance Win32_OperatingSystem
$UptimeDays = [math]::Round(((Get-Date) - $OS.LastBootUpTime).TotalDays, 1)

$InstallDate = if ($OS.InstallDate) { $OS.InstallDate.ToString("yyyy-MM-dd") } else { "Unknown" }

try { $AV = Get-CimInstance -Namespace "root\SecurityCenter2" -Class AntivirusProduct -ErrorAction Stop | Select-Object -ExpandProperty displayName -First 1 } catch { $AV = "Windows Defender / Unknown" }
try { 
    $BitLocker = Get-CimInstance -Namespace "root\CIMv2\Security\MicrosoftVolumeEncryption" -Class Win32_EncryptableVolume -Filter "DriveLetter='C:'" -ErrorAction Stop
    $BLStatus = if ($BitLocker.ProtectionStatus -eq 1) { "Encrypted" } else { "Decrypted/Suspended" }
} catch { $BLStatus = "Not Supported / Off" }

$FwProfiles = Get-NetFirewallProfile | Where-Object Enabled -eq $true | Select-Object -ExpandProperty Name
$FwStatus = if ($FwProfiles) { $FwProfiles -join ", " } else { "All Disabled!" }

$PendingReboot = $false
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $PendingReboot = $true }
if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager") { if ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager").PendingFileRenameOperations) { $PendingReboot = $true } }

$License = Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey is not NULL" | Select-Object -First 1
$Activation = if ($License.LicenseStatus -eq 1) { "Activated" } else { "Not Activated" }

$SecStatus = if ($BLStatus -eq "Encrypted") {"Pass"} else {"Warning"}
$RebootStatus = if ($PendingReboot) {"Warning"} else {"Pass"}

Add-Result "OS & Security" "Windows OS" "<b>Version:</b> $($OS.Caption) ($($OS.OSArchitecture))<br><b>Installed:</b> $InstallDate<br><b>Uptime:</b> $UptimeDays Days<br><b>Activation:</b> $Activation" "OK"
Add-Result "OS & Security" "Security Modules" "<b>Antivirus:</b> $AV<br><b>C: Drive BitLocker:</b> $BLStatus<br><b>Active Firewalls:</b> $FwStatus" $SecStatus
Add-Result "OS & Security" "System State" "<b>Pending Reboot:</b> $($PendingReboot)" $RebootStatus

# 3. CPU & Board
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[2])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Processor & Motherboard..." -ForegroundColor Green
$Board = Get-CimInstance Win32_BaseBoard
Add-Result "Processor & Board" "Baseboard" "<b>MFR:</b> $($Board.Manufacturer)<br><b>Product:</b> $($Board.Product)<br><b>Serial:</b> $($Board.SerialNumber)" "OK"

$CPU = Get-CimInstance Win32_Processor
$CpuStatus = if ($CPU.Status -eq "OK") { "Pass" } else { "Warning" }
$VirtStatus = if ($CPU.VirtualizationFirmwareEnabled) { "Enabled" } else { "Disabled" }
Add-Result "Processor & Board" "Processor (CPU)" "<b>Model:</b> $($CPU.Name)<br><b>Cores/Threads:</b> $($CPU.NumberOfCores)C / $($CPU.NumberOfLogicalProcessors)T<br><b>Base Clock:</b> $($CPU.MaxClockSpeed) MHz<br><b>Virtualization:</b> $VirtStatus<br><b>L2 / L3 Cache:</b> $($CPU.L2CacheSize)KB / $($CPU.L3CacheSize)KB" $CpuStatus

# 4. Memory
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[3])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking RAM & Virtual Memory..." -ForegroundColor Green
$RAMs = Get-CimInstance Win32_PhysicalMemory
$TotalRamGB = [math]::Round(($RAMs | Measure-Object Capacity -Sum).Sum / 1GB, 2)
$MemStatus = if ($TotalRamGB -ge 8) { "Pass" } else { "Low Memory" }

$RamDetails = "<b>Total Capacity:</b> ${TotalRamGB} GB<br>"
foreach ($stick in $RAMs) {
    $Size = [math]::Round($stick.Capacity / 1GB, 1)
    # Replaced bullet points with standard hyphens
    $RamDetails += "- $($stick.DeviceLocator): ${Size}GB $($stick.Manufacturer) @ $($stick.Speed)MHz<br>"
}
Add-Result "Memory & Virtual" "Physical RAM" $RamDetails $MemStatus

$PageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
if ($PageFile) {
    Add-Result "Memory & Virtual" "Page File (Virtual)" "<b>Location:</b> $($PageFile.Name)<br><b>Allocated:</b> $($PageFile.AllocatedBaseSize) MB<br><b>Current Usage:</b> $($PageFile.CurrentUsage) MB" "OK"
}

# 5. Storage
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[4])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Physical Disks & SMART..." -ForegroundColor Green
try {
    $PhysDisks = Get-CimInstance -ClassName MSFT_PhysicalDisk -Namespace root\Microsoft\Windows\Storage
    foreach ($Disk in $PhysDisks) {
        $Type = switch ($Disk.MediaType) { 3 {"HDD"} 4 {"SSD"} Default {"Unknown"} }
        $Bus  = $Disk.BusType
        $SizeGB = [math]::Round($Disk.Size / 1GB, 2)
        $Status = if ($Disk.HealthStatus -eq 0) { "Pass" } else { "Warning" }
        Add-Result "Physical Storage" "Drive: $($Disk.FriendlyName)" "<b>Type:</b> $Type ($Bus)<br><b>Capacity:</b> ${SizeGB} GB<br><b>Operational Status:</b> $($Disk.OperationalStatus)" $Status
    }
} catch {
    $Disks = Get-CimInstance Win32_DiskDrive
    foreach ($Disk in $Disks) {
        $SizeGB = [math]::Round($Disk.Size / 1GB, 2)
        $Status = if ($Disk.Status -eq "OK") { "Pass" } else { "FAIL" }
        Add-Result "Physical Storage" "Drive: $($Disk.Model)" "<b>Capacity:</b> ${SizeGB} GB<br><b>Interface:</b> $($Disk.InterfaceType)<br><b>Partitions:</b> $($Disk.Partitions)" $Status
    }
}

# 6. Logical Volumes
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[5])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Partition Space..." -ForegroundColor Green
$Volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" # Fixed Local Disks Only
foreach ($Vol in $Volumes) {
    if ($Vol.Size -gt 0) {
        $TotalGB = [math]::Round($Vol.Size / 1GB, 2)
        $FreeGB  = [math]::Round($Vol.FreeSpace / 1GB, 2)
        $FreePct = [math]::Round(($Vol.FreeSpace / $Vol.Size) * 100, 1)
        $VolStatus = if ($FreePct -gt 15) { "Pass" } else { "Low Space" }
        Add-Result "Logical Volumes" "Partition: $($Vol.DeviceID)" "<b>Label:</b> $($Vol.VolumeName)<br><b>File System:</b> $($Vol.FileSystem)<br><b>Free Space:</b> $FreeGB GB of $TotalGB GB ($FreePct%)" $VolStatus
    } else {
        Add-Result "Logical Volumes" "Partition: $($Vol.DeviceID)" "<b>Label:</b> $($Vol.VolumeName)<br>Drive reports 0 bytes capacity (Unformatted/Empty)" "Warning"
    }
}

# 7. Graphics & Audio
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[6])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Media Controllers..." -ForegroundColor Green
$GPUs = Get-CimInstance Win32_VideoController
foreach ($GPU in $GPUs) {
    $Status = if ($GPU.Status -eq "OK") { "Pass" } else { "Error" }
    $VRAM = if ($GPU.AdapterRAM) { [math]::Round($GPU.AdapterRAM / 1MB, 0) } else { "Dynamic" }
    $Res = if ($GPU.CurrentHorizontalResolution) { "$($GPU.CurrentHorizontalResolution) x $($GPU.CurrentVerticalResolution) @ $($GPU.CurrentRefreshRate)Hz" } else { "Inactive Display" }
    Add-Result "Graphics & Audio" "GPU: $($GPU.Name)" "<b>Driver:</b> $($GPU.DriverVersion)<br><b>VRAM:</b> $VRAM MB<br><b>Resolution:</b> $Res" $Status
}

$Audio = Get-CimInstance Win32_SoundDevice
$AudioStr = ""
# Replaced bullet points with standard hyphens
foreach ($snd in $Audio) { $AudioStr += "- $($snd.Name)<br>" }
if ($AudioStr) { Add-Result "Graphics & Audio" "Audio Devices" $AudioStr "OK" }

# 8. Network
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[7])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking IPs and Connectivity..." -ForegroundColor Green
$Nets = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
foreach ($Net in $Nets) {
    $IP = $Net.IPAddress[0]
    $GW = if ($Net.DefaultIPGateway) { $Net.DefaultIPGateway[0] } else { "None" }
    Add-Result "Network & Comms" "$($Net.Description)" "<b>IPv4 Address:</b> $IP<br><b>Gateway:</b> $GW<br><b>MAC:</b> $($Net.MACAddress)" "Connected"
}

try {
    $Wifi = netsh wlan show interfaces | Select-String "SSID|Signal|Radio" | Out-String
    if ($Wifi.Trim().Length -gt 0) {
        $WifiClean = $Wifi -replace "`r`n", "<br>" -replace "\s{2,}", " "
        Add-Result "Network & Comms" "Active Wi-Fi Profile" $WifiClean "Pass"
    }
} catch {}

# 9. Hardware Health & Errors
$CurrentTask++; Write-Progress -Activity "Scanning System" -Status "Checking $($DiagnosticsList[8])" -PercentComplete (($CurrentTask/$TotalTasks)*100)
Write-Host "   [$CurrentTask/$TotalTasks] Checking Global Device Manager Errors & Battery..." -ForegroundColor Green

$DeviceErrors = Get-CimInstance Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 -and $_.ConfigManagerErrorCode -ne $null }
if ($DeviceErrors) {
    foreach ($err in $DeviceErrors) { 
        Add-Result "Hardware Health" "Device Failure Detected" "<b>Device:</b> $($err.Name)<br><b>Hardware ID:</b> $($err.DeviceID)<br><b>Error Code:</b> $($err.ConfigManagerErrorCode)" "FAIL" 
    }
} else {
    Add-Result "Hardware Health" "Device Manager Sweep" "No hardware or driver failures detected across any bus." "Pass"
}

try {
    $BatFull = Get-CimInstance -ClassName BatteryFullChargedCapacity -Namespace root\wmi -ErrorAction Stop
    $BatStat = Get-CimInstance -ClassName BatteryStaticData -Namespace root\wmi -ErrorAction Stop
    $Wear = [math]::Round((1 - ($BatFull.FullChargedCapacity / $BatStat.DesignedCapacity)) * 100, 1)
    $WearStatus = if ($Wear -lt 25) { "Pass" } elseif ($Wear -lt 40) { "Warning" } else { "Degraded" }
    Add-Result "Hardware Health" "Internal Battery" "<b>Design Capacity:</b> $($BatStat.DesignedCapacity) mWh<br><b>Full Charge Cap:</b> $($BatFull.FullChargedCapacity) mWh<br><b>Wear Level:</b> $Wear %" $WearStatus
} catch {
    $Battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    if ($Battery) { Add-Result "Hardware Health" "Standard Battery" "<b>Status:</b> $($Battery.Status)<br><b>Charge:</b> $($Battery.EstimatedChargeRemaining)%" "Pass" } 
}

Write-Progress -Activity "Scanning System" -Completed

# --- 6. REPORT GENERATION ---
Write-Host "`n[ REPORT GENERATION ]" -ForegroundColor Yellow
$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$RunTimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$TableRows = ""
$CurrentCategory = ""

foreach ($Row in $ReportItems) {
    if ($Row.Category -ne $CurrentCategory) {
        $TableRows += "<tr class='category-row'><td colspan='3'>$($Row.Category)</td></tr>"
        $CurrentCategory = $Row.Category
    }

    $StatusColor = switch -Regex ($Row.Status) {
        "Pass|OK|Connected|Activated|Encrypted|Enabled" { "#28a745" }
        "Warning|Low Memory|Low Space|Degraded" { "#fd7e14" }
        "FAIL|Error|Not Activated|Disabled" { "#dc3545" }
        Default { "#6c757d" }
    }
    $TableRows += "<tr><td style='width: 25%;'><strong>$($Row.Component)</strong></td><td style='width: 55%;'>$($Row.Details)</td><td style='width: 20%; text-align: center;'><strong style='color:$StatusColor'>$($Row.Status)</strong></td></tr>"
}

$HtmlFile = "$ReportDir\HardwareScan_${TicketNumber}_$CustomerName.html"
$PdfFile  = "$ReportDir\HardwareScan_${TicketNumber}_$CustomerName.pdf"

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; padding: 30px; line-height: 1.4; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0056b3; padding-bottom: 10px; }
    h1 { color: #0056b3; margin-bottom: 5px; font-size: 24px; }
    .meta { font-size: 14px; color: #555; display: flex; justify-content: space-between; margin-top: 15px; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    th { text-align: left; background: #0056b3; color: white; padding: 10px; }
    td { padding: 10px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
    .category-row td { background: #f0f4f8; font-weight: bold; font-size: 14px; color: #0056b3; text-transform: uppercase; border-top: 2px solid #d0dce8; }
    b { color: #222; }
    .footer { text-align: center; font-size: 11px; color: #888; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 10px; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" style="max-height:80px;">
    <h1>Ultimate Hardware & OS Diagnostics Report</h1>
    <div class="meta">
        <div><strong>Ticket:</strong> $TicketNumber<br><strong>Customer:</strong> $CustomerName</div>
        <div style="text-align: right;"><strong>Date:</strong> $CurrentDate<br><strong>Engineer:</strong> $EngineerName</div>
    </div>
</div>

<table>
    <thead><tr><th>Component / Check</th><th>Detailed Specifications & Telemetry</th><th style='text-align: center;'>Health / Status</th></tr></thead>
    <tbody>$TableRows</tbody>
</table>

<div class="footer">
    &copy; $(Get-Date -Format yyyy) by Apollo Technology LTD. Created by Lewis Wiltshire (Apollo Technology).
</div>
</body>
</html>
"@

$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8

# Convert to PDF via Edge
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
        Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Ultimate Hardware Audit: $env:COMPUTERNAME ($TicketNumber)" -Body "Attached is the forensic hardware report for Ticket $TicketNumber ($CustomerName)." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PdfFile -ErrorAction Stop
        Write-Host "   > Email Sent Successfully!" -ForegroundColor Green
    } catch {
        Write-Error "   > Failed to send email. Error: $_"
    }
}

try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Ultimate Diagnostics finished. The report has been opened."

# --- 8. STOP LOGGING ---
if ($VerboseMode) {
    Write-Host "Stopping Verbose Logging..." -ForegroundColor DarkGray
    Stop-Transcript | Out-Null
}

Pause