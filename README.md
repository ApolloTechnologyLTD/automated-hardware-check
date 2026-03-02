# Apollo Technology Hardware & OS Diagnostics

A comprehensive, native PowerShell diagnostic tool designed for IT engineers and support teams. This script performs a deep-dive analysis of a Windows system's hardware, operating system, and security postures, ultimately generating a clean, professional PDF report for your ticketing system or customer records.

## 🚀 Quick Start

Run the script without cloning the repository:

PowerShell

```powershell
iwr https://apollotech.short.gy/automatic-hardware-check -OutFile hardwarecheck.ps1; powershell -ExecutionPolicy Bypass .\hardwarecheck.ps1 -Force
```

> \[!IMPORTANT\] **ADMINISTRATOR PRIVILEGES REQUIRED**
> 
> This script modifies system files/registries. You must launch your terminal with **"Run as Administrator"** rights. If you run this in a standard user shell, the script will fail or behave unexpectedly. *(Note: The script includes an auto-elevation attempt, but launching directly as Admin is best practice).*

---

## ✨ Key Features

This script runs through 9 distinct diagnostic phases, checking the following components:

🔍 **System & Boot:** BIOS version/dates, Secure Boot status, and TPM readiness.

🛡️ **OS & Security:** Windows activation, Uptime, Antivirus status, BitLocker encryption (C: Drive), Firewalls, and Pending Reboot checks.

⚙️ **Processor & Board:** Motherboard serials/models, CPU cores/threads, clock speeds, and Virtualization status.

🧠 **Memory & Virtual:** Physical RAM capacity, individual stick speeds/locations, and Page File usage.

💾 **Physical Storage & Logical Volumes:** HDD/SSD identification, operational health/SMART status, partition labels, and low free-space warnings.

🎮 **Graphics & Audio:** GPU models, VRAM, current resolution, drivers, and connected audio devices.

🌐 **Network & Comms:** Active IPv4 addresses, Gateways, MAC addresses, and saved Wi-Fi profiles.

🩺 **Hardware Health:** Global Device Manager error sweeps (captures failing devices/drivers) and internal laptop battery wear levels.

## 🛠️ How It Works

**Launch**: Upon running the script, it will ensure the console is running with elevated privileges.

**System Locks**: It temporarily disables QuickEdit mode (to prevent accidental console freezing) and prevents the PC from going to sleep during the scan.

**Interactive Prompts**: You will be asked to enter the following information for the final report:

`Engineer Name`

`Ticket Number`

`Customer Name`

**Scanning**: A progress bar will track the 9 diagnostic stages.

**Report Generation**: An HTML report is generated and automatically converted into a PDF using Microsoft Edge's headless mode.

## 📂 Output & Logs

**Default Output Directory:** `C:\temp\Apollo_Reports`

**File Naming Convention:** `HardwareScan_[TicketNumber]_[CustomerName].pdf`

Once completed, the script will automatically open the generated PDF (or HTML if PDF conversion fails) for immediate review.

## ⚙️ Advanced Configuration (For Script Editing)

If you are downloading and modifying the `.ps1` file directly, there are several variables at the top of the script (under `# --- 0. CONFIGURATION ---`) that you can customize:

`$VerboseMode = $true`: Enables hidden background transcript logging to `C:\temp\hwcheck` for debugging.

`$EmailEnabled = $true`: Prompts for a secure password at runtime and automatically emails the finished PDF report to a specified address. Configure `$SmtpServer`, `$FromAddress`, and `$ToAddress` to utilize this feature.

`$ReportDir`: Change the default save location of the reports.