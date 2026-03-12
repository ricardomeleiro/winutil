#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WinUtil - Windows Utility Script
    Served via: irm win.c-innovation.com.br/win.ps1 | iex
.DESCRIPTION
    Comprehensive Windows utility: app install, tweaks, diagnostics, repairs, domain management.
#>

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'

# -------------------------------------------
#  BANNER
# -------------------------------------------
function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  =============================================" -ForegroundColor Cyan
    Write-Host "        W I N U T I L  -  Windows Utility      " -ForegroundColor Cyan
    Write-Host "     irm win.c-innovation.com.br/win.ps1 | iex " -ForegroundColor DarkGray
    Write-Host "  =============================================" -ForegroundColor Cyan
    Write-Host ""
}

# -------------------------------------------
#  HELPERS
# -------------------------------------------
function Write-Status { param($msg, $color = "Green")  Write-Host "  [+] $msg" -ForegroundColor $color }
function Write-Info { param($msg)                     Write-Host "  [i] $msg" -ForegroundColor Yellow }
function Write-Err { param($msg)                     Write-Host "  [!] $msg" -ForegroundColor Red }
function Write-Section { param($msg)                     Write-Host "`n  == $msg ==" -ForegroundColor Magenta }

function Confirm-Action {
    param($prompt)
    $r = Read-Host "  $prompt [Y/N]"
    return ($r -match '^[Yy]')
}

function Check-Winget {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Info "winget not found. Opening Microsoft Store..."
        Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1"
        Read-Host "  Press Enter after winget is installed"
    }
}

# -------------------------------------------
#  0. DEFAULT SOFTWARES
# -------------------------------------------
function Install-DefaultSoftwares {
    $defaults = @(
        @{Name = "Microsoft 365 (Office)"; WinGet = "Microsoft.Office"; Store = $false },
        @{Name = "Microsoft Teams"; WinGet = "Microsoft.Teams"; Store = $false },
        @{Name = "Google Chrome"; WinGet = "Google.Chrome"; Store = $false },
        @{Name = "Lightshot"; WinGet = "Skillbrains.Lightshot"; Store = $false },
        @{Name = "WinRAR"; WinGet = "RARLab.WinRAR"; Store = $false },
        @{Name = "WhatsApp"; WinGet = "9NKSQGP7F2NH"; Store = $true },
        @{Name = "Power BI Desktop"; WinGet = "Microsoft.PowerBI"; Store = $false },
        @{Name = "TeamViewer"; WinGet = "TeamViewer.TeamViewer"; Store = $false },
        @{Name = "LogMeIn"; WinGet = "LogMeIn.LogMeIn"; Store = $false }
    )

    Write-Section "DEFAULT SOFTWARES"
    Write-Host ""
    Write-Host "  [1]  Install All" -ForegroundColor Green

    $i = 2
    foreach ($app in $defaults) {
        Write-Host "  [$i]  $($app.Name)"
        $i++
    }

    Write-Host "  [0]  Cancel" -ForegroundColor Red
    Write-Host ""
    $choice = Read-Host "  Select an option"

    if ($choice -eq '0') { Write-Info "Installation cancelled."; return }

    $toInstall = @()

    if ($choice -eq '1') {
        $toInstall = $defaults
    }
    elseif ($choice -match '^\d+$' -and [int]$choice -ge 2 -and [int]$choice -le ($defaults.Count + 1)) {
        $toInstall = @($defaults[[int]$choice - 2])
    }
    else {
        Write-Err "Invalid option."
        return
    }

    Check-Winget
    Write-Host ""

    $total = $toInstall.Count
    $current = 0

    foreach ($app in $toInstall) {
        $current++
        Write-Host "  [$current/$total] Installing $($app.Name)..." -ForegroundColor Cyan

        if ($app.Store) {
            winget install --id $app.WinGet --source msstore --silent --accept-package-agreements --accept-source-agreements
        }
        else {
            winget install --id $app.WinGet --silent --accept-package-agreements --accept-source-agreements
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Status "$($app.Name) installed successfully." Green
        }
        else {
            Write-Err "$($app.Name) failed or is already installed."
        }
    }

    Write-Host ""
    Write-Status "Done!" Green
}

# -------------------------------------------
#  1. APP INSTALLER
# -------------------------------------------
$Apps = [ordered]@{
    "Browsers"      = @(
        @{Name = "Google Chrome"; WinGet = "Google.Chrome" },
        @{Name = "Mozilla Firefox"; WinGet = "Mozilla.Firefox" },
        @{Name = "Brave"; WinGet = "Brave.Brave" },
        @{Name = "Opera GX"; WinGet = "Opera.OperaGX" }
    )
    "Dev Tools"     = @(
        @{Name = "VS Code"; WinGet = "Microsoft.VisualStudioCode" },
        @{Name = "Git"; WinGet = "Git.Git" },
        @{Name = "Node.js LTS"; WinGet = "OpenJS.NodeJS.LTS" },
        @{Name = "Python 3"; WinGet = "Python.Python.3" },
        @{Name = "Windows Terminal"; WinGet = "Microsoft.WindowsTerminal" },
        @{Name = "Docker Desktop"; WinGet = "Docker.DockerDesktop" }
    )
    "Communication" = @(
        @{Name = "Discord"; WinGet = "Discord.Discord" },
        @{Name = "Slack"; WinGet = "SlackTechnologies.Slack" },
        @{Name = "Microsoft Teams"; WinGet = "Microsoft.Teams" },
        @{Name = "Zoom"; WinGet = "Zoom.Zoom" }
    )
    "Utilities"     = @(
        @{Name = "7-Zip"; WinGet = "7zip.7zip" },
        @{Name = "VLC"; WinGet = "VideoLAN.VLC" },
        @{Name = "Everything"; WinGet = "voidtools.Everything" },
        @{Name = "Notepad++"; WinGet = "Notepad++.Notepad++" },
        @{Name = "PowerToys"; WinGet = "Microsoft.PowerToys" }
    )
    "Security"      = @(
        @{Name = "Malwarebytes"; WinGet = "Malwarebytes.Malwarebytes" },
        @{Name = "Bitwarden"; WinGet = "Bitwarden.Bitwarden" },
        @{Name = "WireGuard"; WinGet = "WireGuard.WireGuard" }
    )
}

function Install-Apps {
    Write-Section "APP INSTALLER"
    Check-Winget
    $toInstall = @()
    $idx = 1
    $indexMap = @{}

    foreach ($category in $Apps.Keys) {
        Write-Host "`n  -- $category --" -ForegroundColor Cyan
        foreach ($app in $Apps[$category]) {
            Write-Host "    [$idx] $($app.Name)"
            $indexMap["$idx"] = $app
            $idx++
        }
    }

    Write-Host ""
    $selection = Read-Host "  Enter numbers to install (comma-separated, or 'all')"

    if ($selection -eq 'all') {
        foreach ($cat in $Apps.Keys) { $toInstall += $Apps[$cat] }
    }
    else {
        foreach ($num in ($selection -split ',')) {
            $num = $num.Trim()
            if ($indexMap.ContainsKey($num)) { $toInstall += $indexMap[$num] }
        }
    }

    if ($toInstall.Count -eq 0) { Write-Err "No apps selected."; return }

    Write-Host ""
    foreach ($app in $toInstall) {
        Write-Status "Installing $($app.Name)..."
        winget install --id $app.WinGet --silent --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -eq 0) {
            Write-Status "$($app.Name) installed." Green
        }
        else {
            Write-Err "$($app.Name) failed or already installed."
        }
    }
}

# -------------------------------------------
#  2. WINDOWS TWEAKS
# -------------------------------------------
function Apply-Tweaks {
    Write-Section "WINDOWS TWEAKS"
    Write-Host ""
    Write-Host "  [1] Privacy     - disable telemetry, ads, tracking"
    Write-Host "  [2] Performance - power plan, visual effects, startup"
    Write-Host "  [3] UI and UX   - dark mode, taskbar, explorer"
    Write-Host "  [4] All of the above"
    Write-Host ""
    $choice = Read-Host "  Select"

    if ($choice -in '1', '4') {
        Write-Status "Applying Privacy Tweaks..."
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0 -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableActivityFeed" -Value 0 -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "PublishUserActivities" -Value 0 -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Siuf\Rules" -Name "NumberOfSIUFInPeriod" -Value 0 -Force
        Write-Status "Privacy tweaks applied." Green
    }

    if ($choice -in '2', '4') {
        Write-Status "Applying Performance Tweaks..."
        powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Value 2 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" -Name "StartupDelayInMSec" -Value 0 -Force
        powercfg /hibernate off
        Stop-Service -Name SysMain -Force
        Set-Service -Name SysMain -StartupType Disabled
        Write-Status "Performance tweaks applied." Green
    }

    if ($choice -in '3', '4') {
        Write-Status "Applying UI/UX Tweaks..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 1 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0 -Force
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Value 0 -Force
        Write-Status "UI/UX tweaks applied." Green
    }
}

# -------------------------------------------
#  3. SYSTEM INFO / DIAGNOSTICS
# -------------------------------------------
function Show-SystemInfo {
    Write-Section "SYSTEM DIAGNOSTICS"
    Write-Host ""

    $os = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor
    $ram = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
    $disk = Get-PSDrive C
    $gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
    $net = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
    $ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' } | Select-Object -First 1).IPAddress

    $ramGB = [math]::Round($ram.Sum / 1GB, 1)
    $diskFree = [math]::Round($disk.Free / 1GB, 1)
    $diskUsed = [math]::Round($disk.Used / 1GB, 1)

    Write-Host "  +---------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  | OS        : $($os.Caption)" -ForegroundColor White
    Write-Host "  | Build     : $($os.BuildNumber)  |  Arch: $($os.OSArchitecture)" -ForegroundColor White
    Write-Host "  | CPU       : $($cpu.Name)" -ForegroundColor White
    Write-Host "  | Cores     : $($cpu.NumberOfCores) Physical / $($cpu.NumberOfLogicalProcessors) Logical" -ForegroundColor White
    Write-Host "  | RAM       : ${ramGB} GB" -ForegroundColor White
    Write-Host "  | GPU       : $($gpu.Name)" -ForegroundColor White
    Write-Host "  | Disk C:\  : ${diskUsed} GB used / ${diskFree} GB free" -ForegroundColor White
    Write-Host "  | Network   : $($net.Name) - $ip" -ForegroundColor White
    Write-Host "  | Hostname  : $env:COMPUTERNAME" -ForegroundColor White
    Write-Host "  | User      : $env:USERNAME" -ForegroundColor White
    Write-Host "  | Uptime    : $(((Get-Date) - $os.LastBootUpTime).ToString('d\d\ h\h\ m\m'))" -ForegroundColor White
    Write-Host "  +---------------------------------------------+" -ForegroundColor DarkCyan

    Write-Host ""
    Write-Info "Top 5 CPU-consuming processes:"
    Get-Process | Sort-Object CPU -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host ("  {0,-30} CPU: {1,8:F1}s  RAM: {2,6} MB" -f $_.Name, $_.CPU, [math]::Round($_.WorkingSet64 / 1MB, 1)) -ForegroundColor Gray
    }

    Write-Host ""
    Write-Info "Checking Windows Defender status..."
    $defender = Get-MpComputerStatus
    Write-Host "  Antivirus enabled : $($defender.AntivirusEnabled)" -ForegroundColor $(if ($defender.AntivirusEnabled) { "Green" }else { "Red" })
    Write-Host "  Definitions date  : $($defender.AntivirusSignatureLastUpdated)" -ForegroundColor Gray
}

# -------------------------------------------
#  4. FIX AND REPAIR TOOLS
# -------------------------------------------
function Run-Repairs {
    Write-Section "FIX AND REPAIR TOOLS"
    Write-Host ""
    Write-Host "  [1] SFC + DISM  (System File Checker and Image Repair)"
    Write-Host "  [2] Fix Windows Update"
    Write-Host "  [3] Flush DNS + Reset Network Stack"
    Write-Host "  [4] Clear Temp Files and Disk Cleanup"
    Write-Host "  [5] Reset Windows Store / App Cache"
    Write-Host "  [6] Run All Repairs"
    Write-Host ""
    $choice = Read-Host "  Select"

    if ($choice -in '1', '6') {
        Write-Status "Running SFC..."
        sfc /scannow
        Write-Status "Running DISM RestoreHealth..."
        DISM /Online /Cleanup-Image /RestoreHealth
        Write-Status "SFC + DISM complete." Green
    }

    if ($choice -in '2', '6') {
        Write-Status "Fixing Windows Update..."
        Stop-Service -Name wuauserv, cryptSvc, bits, msiserver -Force
        Remove-Item "$env:SystemRoot\SoftwareDistribution" -Recurse -Force
        Remove-Item "$env:SystemRoot\System32\catroot2" -Recurse -Force
        Start-Service -Name wuauserv, cryptSvc, bits, msiserver
        Write-Status "Windows Update cache cleared and services restarted." Green
    }

    if ($choice -in '3', '6') {
        Write-Status "Flushing DNS..."
        ipconfig /flushdns
        ipconfig /registerdns
        Write-Status "Resetting network stack..."
        netsh int ip reset
        netsh winsock reset
        netsh advfirewall reset
        Write-Status "Network stack reset. A reboot is recommended." Yellow
    }

    if ($choice -in '4', '6') {
        Write-Status "Clearing Temp files..."
        Remove-Item "$env:TEMP\*" -Recurse -Force
        Remove-Item "C:\Windows\Temp\*" -Recurse -Force
        Remove-Item "C:\Windows\Prefetch\*" -Recurse -Force
        cleanmgr /sagerun:1
        Write-Status "Temp files cleared." Green
    }

    if ($choice -in '5', '6') {
        Write-Status "Resetting Windows Store cache..."
        wsreset.exe
        Write-Status "Store cache reset." Green
    }
}

# -------------------------------------------
#  5. DOMAIN MANAGEMENT
# -------------------------------------------
function Join-DomainHelper {
    param($domainName)
    $prefix = $domainName.Split('.')[0].ToUpper()
    $user = Read-Host "  Admin username (e.g. $prefix\Administrator)"
    $pass = Read-Host "  Password" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($user, $pass)
    try {
        Add-Computer -DomainName $domainName -Credential $cred -Force
        Write-Status "Successfully joined domain '$domainName'. Please reboot." Green
        if (Confirm-Action "Reboot now?") { Restart-Computer -Force }
    }
    catch {
        Write-Err "Failed to join domain '$domainName': $_"
    }
}

function Manage-Domain {
    Write-Section "DOMAIN MANAGEMENT"
    Write-Host ""

    $currentDomain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomain = (Get-WmiObject Win32_ComputerSystem).PartOfDomain
    Write-Info "Current: $(if($isDomain){'Domain: ' + $currentDomain}else{'Workgroup: ' + $currentDomain})"

    Write-Host ""
    Write-Host "  -- Join a preset domain --" -ForegroundColor DarkCyan
    Write-Host "  [1]  Join ONSHORE  domain  (chouest-br.local)" -ForegroundColor Cyan
    Write-Host "  [2]  Join OFFSHORE domain  (rov.local)"        -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  -- Other options --" -ForegroundColor DarkCyan
    Write-Host "  [3]  Join a custom domain"
    Write-Host "  [4]  Leave domain (join Workgroup)"
    Write-Host "  [5]  Show domain / AD info"
    Write-Host ""
    $choice = Read-Host "  Select"

    if ($choice -eq '1') {
        Write-Info "Joining ONSHORE domain: chouest-br.local"
        Join-DomainHelper "chouest-br.local"
    }

    if ($choice -eq '2') {
        Write-Info "Joining OFFSHORE domain: rov.local"
        Join-DomainHelper "rov.local"
    }

    if ($choice -eq '3') {
        $domain = Read-Host "  Enter domain name (e.g. corp.contoso.com)"
        Join-DomainHelper $domain
    }

    if ($choice -eq '4') {
        $wg = Read-Host "  Enter workgroup name (default: WORKGROUP)"
        if ([string]::IsNullOrWhiteSpace($wg)) { $wg = "WORKGROUP" }
        $user = Read-Host "  Current domain admin username"
        $pass = Read-Host "  Password" -AsSecureString
        $cred = New-Object System.Management.Automation.PSCredential($user, $pass)
        try {
            Remove-Computer -WorkgroupName $wg -Credential $cred -Force
            Write-Status "Left domain. Joined workgroup '$wg'. Please reboot." Green
            if (Confirm-Action "Reboot now?") { Restart-Computer -Force }
        }
        catch {
            Write-Err "Failed to leave domain: $_"
        }
    }

    if ($choice -eq '5') {
        Write-Host ""
        $cs = Get-WmiObject Win32_ComputerSystem
        Write-Host "  Computer Name  : $($cs.Name)"        -ForegroundColor White
        Write-Host "  Domain / WG    : $($cs.Domain)"       -ForegroundColor White
        Write-Host "  Part of Domain : $($cs.PartOfDomain)" -ForegroundColor White
        if ($isDomain) {
            $dc = (nltest /dsgetdc:$currentDomain 2>&1 | Select-String "DC:").ToString().Trim()
            Write-Host "  Domain Controller: $dc" -ForegroundColor White
            Write-Host ""
            Write-Info "Local users:"
            Get-LocalUser | ForEach-Object {
                Write-Host "   - $($_.Name)  [Enabled: $($_.Enabled)]" -ForegroundColor Gray
            }
        }
    }
}

# -------------------------------------------
#  MAIN MENU
# -------------------------------------------
function Show-Menu {
    Write-Host "  +------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  |  [1]  Default Softwares            |" -ForegroundColor Green
    Write-Host "  |  [2]  App Installer                |" -ForegroundColor White
    Write-Host "  |  [3]  Windows Tweaks               |" -ForegroundColor White
    Write-Host "  |  [4]  System Info / Diagnostics    |" -ForegroundColor White
    Write-Host "  |  [5]  Fix and Repair Tools         |" -ForegroundColor White
    Write-Host "  |  [6]  Domain Management            |" -ForegroundColor White
    Write-Host "  |  [Q]  Quit                         |" -ForegroundColor White
    Write-Host "  +------------------------------------+" -ForegroundColor DarkCyan
    Write-Host ""
}

# -------------------------------------------
#  ENTRY POINT
# -------------------------------------------
Show-Banner

do {
    Show-Menu
    $selection = Read-Host "  Select an option"
    switch ($selection.ToUpper()) {
        '1' { Install-DefaultSoftwares }
        '2' { Install-Apps }
        '3' { Apply-Tweaks }
        '4' { Show-SystemInfo }
        '5' { Run-Repairs }
        '6' { Manage-Domain }
        'Q' {
            Write-Host ""
            Write-Host "  Goodbye!" -ForegroundColor Cyan
            Write-Host ""
            exit
        }
        default { Write-Err "Invalid option. Please select 1-5 or Q." }
    }
    Write-Host ""
    Read-Host "  Press Enter to return to menu"
    Show-Banner
} while ($true)
