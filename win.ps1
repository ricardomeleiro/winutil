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

function Install-WingetLatest {
    Write-Info "Downloading and installing the latest winget (App Installer)..."
    try {
        $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
        $msixBundle = $releases.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1
        $vcLibs = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        $uiXaml  = ($releases.assets | Where-Object { $_.name -like "*.appx" -and $_.name -like "*UIXaml*" } | Select-Object -First 1)

        $tmpDir = "$env:TEMP\winget-install"
        New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

        Write-Info "Downloading VCLibs..."
        Invoke-WebRequest -Uri $vcLibs -OutFile "$tmpDir\vclibs.appx" -UseBasicParsing

        if ($uiXaml) {
            Write-Info "Downloading UI.Xaml..."
            Invoke-WebRequest -Uri $uiXaml.browser_download_url -OutFile "$tmpDir\uixaml.appx" -UseBasicParsing
            Add-AppxPackage -Path "$tmpDir\uixaml.appx" -ErrorAction SilentlyContinue
        }

        Add-AppxPackage -Path "$tmpDir\vclibs.appx" -ErrorAction SilentlyContinue

        Write-Info "Downloading winget $($releases.tag_name)..."
        Invoke-WebRequest -Uri $msixBundle.browser_download_url -OutFile "$tmpDir\winget.msixbundle" -UseBasicParsing
        Add-AppxPackage -Path "$tmpDir\winget.msixbundle"

        Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
        Write-Status "Winget installed successfully." Green
    } catch {
        Write-Warning "Automatic install failed: $_"
        Write-Info "Opening Microsoft Store as fallback..."
        Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1"
        Read-Host "  Press Enter after winget is installed"
    }
}

function Check-Winget {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Info "winget not found. Installing automatically..."
        Install-WingetLatest
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            Write-Warning "winget still not available after install. Please restart the script."
            exit
        }
    }

    # Check if winget version is outdated (minimum v1.6)
    $wingetVersion = (winget --version) -replace 'v',''
    $major = [int]($wingetVersion.Split('.')[0])
    $minor = [int]($wingetVersion.Split('.')[1])

    if ($major -lt 1 -or ($major -eq 1 -and $minor -lt 6)) {
        Write-Info "Winget version $wingetVersion is outdated. Updating automatically..."
        Install-WingetLatest
        $wingetVersion = (winget --version) -replace 'v',''
    }

    Write-Status "Winget v$wingetVersion detected." Green
}

# -------------------------------------------
#  0. DEFAULT SOFTWARES
# -------------------------------------------
function Fix-WingetCertificate {
    Write-Info "Repairing winget sources..."

    # Step 1 - Delete corrupted winget source cache
    $localAppData = [System.Environment]::GetFolderPath("LocalApplicationData")
    $wingetCache = "$localAppData\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\Microsoft.Winget.Source_8wekyb3d8bbwe"
    $wingetCache2 = "$localAppData\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\Microsoft.Winget.MSStore.Source_8wekyb3d8bbwe"

    if (Test-Path $wingetCache)  { Remove-Item $wingetCache  -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $wingetCache2) { Remove-Item $wingetCache2 -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Info "Winget cache cleared."

    # Step 2 - Bypass cert and update winget itself
    winget settings --enable BypassCertificatePinningForMicrosoftStore
    winget upgrade Microsoft.AppInstaller --accept-source-agreements --accept-package-agreements
    winget settings --disable BypassCertificatePinningForMicrosoftStore

    # Step 3 - Remove and re-add sources from scratch
    winget source remove winget   2>&1 | Out-Null
    winget source remove msstore  2>&1 | Out-Null
    winget source add winget  https://cdn.winget.microsoft.com/cache --accept-source-agreements 2>&1 | Out-Null
    winget source add msstore https://storeedgefd.dsx.mp.microsoft.com/v9.0 --type Microsoft.Rest --accept-source-agreements 2>&1 | Out-Null

    # Step 4 - Force update all sources
    winget source update --accept-source-agreements

    # Step 5 - Accept msstore terms
    winget search --source msstore --accept-source-agreements --query "dummy" 2>&1 | Out-Null

    Write-Status "Winget sources repaired and ready." Green
}
 
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

    if (-not (Confirm-Action "Proceed with installation?")) {
        Write-Info "Installation cancelled."
        return
    }

    Check-Winget
    Write-Host ""

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

    $total = $defaults.Count
    $current = 0

    foreach ($app in $defaults) {
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
    Write-Status "Default software installation complete!" Green
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
#  7. AD MANAGEMENT
# -------------------------------------------

# AD connection variables (set when user enters this section)
$adSession = $null
$adDomain = "CHOUEST-BR.local"
$adBaseDN = "DC=RIO-ADDS,DC=chouest-br,DC=local"
$adDC = "10.131.0.11"
$adConnected = $false

function Connect-ADSession {
    if ($script:adConnected -and $script:adSession) { return $true }

    Write-Host ""
    Write-Host "  ==============================================" -ForegroundColor Cyan
    Write-Host "     AUTENTICAÇÃO NO CONTROLADOR DE DOMÍNIO    " -ForegroundColor Yellow
    Write-Host "  ==============================================" -ForegroundColor Cyan

    $adminUser = "Administrator"
    $adminPass = Read-Host "  Digite a senha do Administrador ($adminUser)" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($adminUser, $adminPass)

    try {
        $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $script:adDC -Force -ErrorAction Stop

        $script:adSession = New-PSSession -ComputerName $script:adDC -Credential $cred -ErrorAction Stop
        Invoke-Command -Session $script:adSession -ScriptBlock { Import-Module ActiveDirectory } -ErrorAction Stop
        $script:adConnected = $true
        Write-Status "Conectado ao DC $($script:adDC) com sucesso." Green
        return $true
    }
    catch {
        Write-Err "Erro ao conectar ao DC: $_"
        return $false
    }
}

function Disconnect-ADSession {
    if ($script:adSession) {
        Remove-PSSession $script:adSession
        $script:adSession = $null
        $script:adConnected = $false
        Write-Info "Sessão AD encerrada."
    }
}

function Get-ADOUPath {
    param ([string]$baseOUChoice, [string]$subOUName)
    switch ($baseOUChoice) {
        "1" { return "OU=$subOUName,OU=Departamentos,$($script:adBaseDN)" }
        "2" { return "OU=$subOUName,OU=Consultoria,$($script:adBaseDN)" }
        "3" { return "OU=Usuarios,OU=$subOUName,OU=Filiais,$($script:adBaseDN)" }
        "4" { return "OU=$subOUName,OU=Usuarios de Servicos de TI,$($script:adBaseDN)" }
        default { Write-Err "Opção de OU inválida."; return $null }
    }
}

function Select-ADOU {
    Write-Host ""
    Write-Host "  Escolha a OU base:" -ForegroundColor Cyan
    Write-Host "  [1] Departamentos"
    Write-Host "  [2] Consultoria"
    Write-Host "  [3] Filiais"
    Write-Host "  [4] Usuarios de Servicos de TI"
    $baseChoice = Read-Host "  Selecione"
    switch ($baseChoice) {
        "1" { $sub = Read-Host "  Setor (ex: T.I, RH, Financeiro)" }
        "2" { $sub = Read-Host "  Empresa (ex: EmpresaX)" }
        "3" { $sub = Read-Host "  Filial (ex: Bahia, Betim)" }
        "4" { $sub = Read-Host "  Sub-OU (ex: BI, SAP, RM)" }
        default { Write-Err "Opção inválida."; return $null }
    }
    return Get-ADOUPath -baseOUChoice $baseChoice -subOUName $sub
}

function AD-CreateUser {
    $firstName = Read-Host "  Primeiro nome"
    $lastName = Read-Host "  Sobrenome"
    $cargo = Read-Host "  Cargo"
    $email = Read-Host "  E-mail"
    $username = Read-Host "  Login (SamAccountName)"
    $password = Read-Host "  Senha" -AsSecureString
    $description = Read-Host "  Descrição"

    try {
        $allOUs = Invoke-Command -Session $script:adSession -ScriptBlock {
            Get-ADOrganizationalUnit -Filter * -Properties CanonicalName |
            Sort-Object CanonicalName | Select-Object Name, DistinguishedName, CanonicalName
        }
        if (-not $allOUs) { Write-Err "Nenhuma OU encontrada."; return }

        Write-Host ""
        $i = 1; $ouMap = @{}
        foreach ($ou in $allOUs) {
            Write-Host "  [$i] $($ou.CanonicalName)" -ForegroundColor Gray
            $ouMap[$i] = @{ DN = $ou.DistinguishedName }
            $i++
        }
        Write-Host "  [0] Cancelar"
        $sel = Read-Host "  Selecione a OU"
        if ($sel -eq '0') { return }
        if (-not $ouMap.ContainsKey([int]$sel)) { Write-Err "Opção inválida."; return }

        $ouPath = $ouMap[[int]$sel].DN

        $result = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($fn, $ln, $cargo, $desc, $email, $user, $pass, $ou, $domain)
            if (Get-ADUser -Filter { SamAccountName -eq $user } -ErrorAction SilentlyContinue) {
                throw "Usuário '$user' já existe."
            }
            New-ADUser -Name "$fn $ln" -DisplayName "$fn $ln" -GivenName $fn -Surname $ln `
                -Title $cargo -Description $desc -EmailAddress $email `
                -SamAccountName $user -UserPrincipalName "$user@$domain" `
                -AccountPassword $pass -Enabled $true -Path $ou -ErrorAction Stop
            (Get-ADUser $user -Properties CanonicalName).CanonicalName
        } -ArgumentList $firstName, $lastName, $cargo, $description, $email, $username, $password, $ouPath, $script:adDomain

        Write-Status "Usuário $username criado em: $result" Green
    }
    catch { Write-Err "Erro: $_" }
}

function AD-DisableUser { $u = Read-Host "  Usuário"; try { Invoke-Command -Session $script:adSession -ScriptBlock { param($u) Disable-ADAccount -Identity $u -ErrorAction Stop } -ArgumentList $u; Write-Status "Usuário $u inativado." Green } catch { Write-Err "$_" } }
function AD-EnableUser { $u = Read-Host "  Usuário"; try { Invoke-Command -Session $script:adSession -ScriptBlock { param($u) Enable-ADAccount -Identity $u -ErrorAction Stop } -ArgumentList $u; Write-Status "Usuário $u reativado." Green } catch { Write-Err "$_" } }
function AD-DeleteUser { $u = Read-Host "  Usuário"; try { Invoke-Command -Session $script:adSession -ScriptBlock { param($u) Remove-ADUser -Identity $u -Confirm:$false -ErrorAction Stop } -ArgumentList $u; Write-Status "Usuário $u deletado." Green } catch { Write-Err "$_" } }
function AD-UnlockUser { $u = Read-Host "  Usuário"; try { Invoke-Command -Session $script:adSession -ScriptBlock { param($u) Unlock-ADAccount -Identity $u -ErrorAction Stop } -ArgumentList $u; Write-Status "Usuário $u desbloqueado." Green } catch { Write-Err "$_" } }

function AD-ResetPassword {
    $u = Read-Host "  Usuário"
    $p = Read-Host "  Nova senha" -AsSecureString
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($u, $p) Set-ADAccountPassword -Identity $u -NewPassword $p -Reset -ErrorAction Stop
        } -ArgumentList $u, $p
        Write-Status "Senha de $u resetada." Green
    }
    catch { Write-Err "$_" }
}

function AD-ChangeExtension {
    $u = Read-Host "  Usuário"
    $e = Read-Host "  Novo ramal"
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($u, $e) Set-ADUser -Identity $u -OfficePhone $e -ErrorAction Stop
        } -ArgumentList $u, $e
        Write-Status "Ramal de $u alterado para $e." Green
    }
    catch { Write-Err "$_" }
}

function AD-AddComputer {
    $c = Read-Host "  Nome do computador"
    $ouPath = Select-ADOU
    if ($ouPath) {
        try {
            Invoke-Command -Session $script:adSession -ScriptBlock {
                param($c, $p) New-ADComputer -Name $c -Path $p -ErrorAction Stop
            } -ArgumentList $c, $ouPath
            Write-Status "Computador $c adicionado." Green
        }
        catch { Write-Err "$_" }
    }
}

function AD-RemoveComputer {
    $c = Read-Host "  Nome do computador"
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($c) Remove-ADComputer -Identity $c -Confirm:$false -ErrorAction Stop
        } -ArgumentList $c
        Write-Status "Computador $c removido." Green
    }
    catch { Write-Err "$_" }
}

function AD-MoveObject {
    $obj = Read-Host "  Nome do objeto (usuário ou computador)"
    $newOU = Select-ADOU
    if ($newOU) {
        try {
            Invoke-Command -Session $script:adSession -ScriptBlock {
                param($o, $t) Get-ADObject -Filter { Name -eq $o } | Move-ADObject -TargetPath $t -ErrorAction Stop
            } -ArgumentList $obj, $newOU
            Write-Status "Objeto $obj movido com sucesso." Green
        }
        catch { Write-Err "$_" }
    }
}

function AD-AddUserToGroup {
    $u = Read-Host "  Usuário"; $g = Read-Host "  Grupo"
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($u, $g) Add-ADGroupMember -Identity $g -Members $u -ErrorAction Stop
        } -ArgumentList $u, $g
        Write-Status "Usuário $u adicionado ao grupo $g." Green
    }
    catch { Write-Err "$_" }
}

function AD-RemoveUserFromGroup {
    $u = Read-Host "  Usuário"; $g = Read-Host "  Grupo"
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($u, $g) Remove-ADGroupMember -Identity $g -Members $u -Confirm:$false -ErrorAction Stop
        } -ArgumentList $u, $g
        Write-Status "Usuário $u removido do grupo $g." Green
    }
    catch { Write-Err "$_" }
}

function AD-GetGroupMembers {
    $g = Read-Host "  Grupo"
    try {
        $members = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($g) Get-ADGroupMember -Identity $g | Select-Object Name, SamAccountName
        } -ArgumentList $g
        $members | Format-Table -AutoSize
    }
    catch { Write-Err "$_" }
}

function AD-SetUserAttribute {
    $u = Read-Host "  Usuário"; $a = Read-Host "  Atributo (ex: Title, Department)"; $v = Read-Host "  Novo valor"
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($u, $a, $v) Set-ADUser -Identity $u -Replace @{ $a = $v } -ErrorAction Stop
        } -ArgumentList $u, $a, $v
        Write-Status "Atributo $a de $u alterado para $v." Green
    }
    catch { Write-Err "$_" }
}

function AD-ListUsers {
    $ouPath = Select-ADOU
    if ($ouPath) {
        try {
            $users = Invoke-Command -Session $script:adSession -ScriptBlock {
                param($p) Get-ADUser -Filter * -SearchBase $p -Properties Name, SamAccountName, Enabled |
                Select-Object Name, SamAccountName, Enabled
            } -ArgumentList $ouPath
            $users | Format-Table -AutoSize
        }
        catch { Write-Err "$_" }
    }
}

function AD-ListComputers {
    $ouPath = Select-ADOU
    if ($ouPath) {
        try {
            $computers = Invoke-Command -Session $script:adSession -ScriptBlock {
                param($p) Get-ADComputer -Filter * -SearchBase $p -Properties Name, Enabled |
                Select-Object Name, Enabled
            } -ArgumentList $ouPath
            $computers | Format-Table -AutoSize
        }
        catch { Write-Err "$_" }
    }
}

function AD-SyncAD {
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        } -ErrorAction Stop
        Write-Status "Sincronização AD concluída." Green
    }
    catch { Write-Err "Erro ao sincronizar AD: $_" }
}

function AD-ExportUsers {
    $ouPath = Select-ADOU
    $file = Read-Host "  Arquivo de saída (ex: usuarios.csv)"
    if ($ouPath) {
        try {
            $users = Invoke-Command -Session $script:adSession -ScriptBlock {
                param($p) Get-ADUser -Filter * -SearchBase $p -Properties Name, SamAccountName, Enabled |
                Select-Object Name, SamAccountName, Enabled
            } -ArgumentList $ouPath
            $users | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
            Write-Status "Exportado para $file." Green
        }
        catch { Write-Err "$_" }
    }
}

function AD-ExportComputers {
    $ouPath = Select-ADOU
    $file = Read-Host "  Arquivo de saída (ex: computadores.csv)"
    if ($ouPath) {
        try {
            $computers = Invoke-Command -Session $script:adSession -ScriptBlock {
                param($p) Get-ADComputer -Filter * -SearchBase $p -Properties Name, Enabled, LastLogonDate |
                Select-Object Name, Enabled, LastLogonDate
            } -ArgumentList $ouPath
            $computers | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
            Write-Status "Exportado para $file." Green
        }
        catch { Write-Err "$_" }
    }
}

function AD-ExportGroups {
    $setor = Read-Host "  Setor (ex: T.I, RH)"
    $ouPath = "OU=Grupos,OU=$setor,OU=Departamentos,$($script:adBaseDN)"
    $file = Read-Host "  Arquivo de saída (ex: grupos.csv)"
    try {
        $groups = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($p) Get-ADGroup -Filter * -SearchBase $p -Properties Name, SamAccountName, GroupCategory, GroupScope |
            Select-Object Name, SamAccountName, GroupCategory, GroupScope
        } -ArgumentList $ouPath
        $groups | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        Write-Status "Exportado para $file." Green
    }
    catch { Write-Err "$_" }
}

function Manage-AD {
    if (-not (Connect-ADSession)) { return }

    do {
        Write-Host ""
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host "         GERENCIAMENTO DO ACTIVE DIRECTORY      " -ForegroundColor Yellow
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host "  Domínio: $adDomain  |  DC: $adDC" -ForegroundColor Green
        Write-Host "  ----------------------------------------------"
        Write-Host "  -- Usuários --" -ForegroundColor DarkCyan
        Write-Host "   [1]  Criar usuário"
        Write-Host "   [2]  Inativar usuário"
        Write-Host "   [3]  Reativar usuário"
        Write-Host "   [4]  Deletar usuário"
        Write-Host "   [5]  Resetar senha"
        Write-Host "   [6]  Desbloquear usuário"
        Write-Host "   [7]  Alterar ramal"
        Write-Host "   [8]  Alterar atributo"
        Write-Host "   [9]  Listar usuários"
        Write-Host "  -- Computadores --" -ForegroundColor DarkCyan
        Write-Host "  [10]  Adicionar computador"
        Write-Host "  [11]  Remover computador"
        Write-Host "  [12]  Listar computadores"
        Write-Host "  -- Grupos --" -ForegroundColor DarkCyan
        Write-Host "  [13]  Adicionar usuário a grupo"
        Write-Host "  [14]  Remover usuário de grupo"
        Write-Host "  [15]  Ver membros de grupo"
        Write-Host "  -- Outros --" -ForegroundColor DarkCyan
        Write-Host "  [16]  Mover objeto para outra OU"
        Write-Host "  [17]  Sincronizar AD"
        Write-Host "  [18]  Exportar relatório de usuários"
        Write-Host "  [19]  Exportar relatório de computadores"
        Write-Host "  [20]  Exportar relatório de grupos"
        Write-Host "   [0]  Voltar ao menu principal" -ForegroundColor Red
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host ""

        $adChoice = Read-Host "  Selecione"
        switch ($adChoice) {
            '1' { AD-CreateUser }
            '2' { AD-DisableUser }
            '3' { AD-EnableUser }
            '4' { AD-DeleteUser }
            '5' { AD-ResetPassword }
            '6' { AD-UnlockUser }
            '7' { AD-ChangeExtension }
            '8' { AD-SetUserAttribute }
            '9' { AD-ListUsers }
            '10' { AD-AddComputer }
            '11' { AD-RemoveComputer }
            '12' { AD-ListComputers }
            '13' { AD-AddUserToGroup }
            '14' { AD-RemoveUserFromGroup }
            '15' { AD-GetGroupMembers }
            '16' { AD-MoveObject }
            '17' { AD-SyncAD }
            '18' { AD-ExportUsers }
            '19' { AD-ExportComputers }
            '20' { AD-ExportGroups }
            '0' { Disconnect-ADSession; break }
            default { Write-Err "Opção inválida." }
        }

        if ($adChoice -ne '0') {
            Write-Host ""
            Read-Host "  Pressione Enter para continuar"
        }

    } while ($adChoice -ne '0')
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
    Write-Host "  |  [7]  AD Management                |" -ForegroundColor Cyan
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
        '7' { Manage-AD }
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
