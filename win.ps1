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
        $uiXaml = ($releases.assets | Where-Object { $_.name -like "*.appx" -and $_.name -like "*UIXaml*" } | Select-Object -First 1)

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
    }
    catch {
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
    $wingetVersion = (winget --version) -replace 'v', ''
    $major = [int]($wingetVersion.Split('.')[0])
    $minor = [int]($wingetVersion.Split('.')[1])

    if ($major -lt 1 -or ($major -eq 1 -and $minor -lt 6)) {
        Write-Info "Winget version $wingetVersion is outdated. Updating automatically..."
        Install-WingetLatest
        $wingetVersion = (winget --version) -replace 'v', ''
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

    if (Test-Path $wingetCache) { Remove-Item $wingetCache  -Recurse -Force -ErrorAction SilentlyContinue }
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
        @{Name = "Adobe Acrobat Reader"; WinGet = "Adobe.Acrobat.Reader.64-bit"; Store = $false },
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
        @{Name = "Adobe Acrobat Reader"; WinGet = "Adobe.Acrobat.Reader.64-bit"; Store = $false },
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
        @{Name = "WireGuard"; WinGet = "WireGuard.WireGuard" },
        @{Name = "Wazuh Agent"; WinGet = "Wazuh.WazuhAgent" }
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

    Write-Host ""

    # 1. Configurar DNS apontando para o DC (necessário para resolver o domínio)
    $dcIP = Read-Host "  IP do Controlador de Domínio (necessário para DNS)"
    if (-not [string]::IsNullOrWhiteSpace($dcIP)) {
        try {
            $adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
            if ($adapter) {
                Set-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -ServerAddresses $dcIP -ErrorAction Stop
                Write-Status "DNS configurado para $dcIP na interface '$($adapter.Name)'." Green
                Start-Sleep -Seconds 3
            }
        }
        catch { Write-Info "Não foi possível configurar DNS automaticamente: $_" }
    }

    # 2. Verificar conectividade com o domínio
    Write-Info "Verificando conectividade com '$domainName'..."
    $reachable = Test-Connection -ComputerName $domainName -Count 2 -Quiet -ErrorAction SilentlyContinue
    if (-not $reachable) {
        Write-Err "Não foi possível resolver/alcançar '$domainName'."
        Write-Info "Verifique se o DNS está correto e se há rota de rede para o DC."
        if (-not (Confirm-Action "Tentar ingressar mesmo assim?")) { return }
    }

    # 3. Credenciais
    $user = Read-Host "  Usuário admin do domínio (ex: $prefix\Administrator)"
    $pass = Read-Host "  Senha" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($user, $pass)

    # 4. OU de destino (opcional)
    $ouPath = Read-Host "  OU de destino (deixe vazio para usar 'Computers' padrão)"

    try {
        $params = @{
            DomainName  = $domainName
            Credential  = $cred
            Force       = $true
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($ouPath)) { $params['OUPath'] = $ouPath }

        Write-Info "Ingressando no domínio '$domainName'..."
        Add-Computer @params

        # 5. Verificar se o ingresso foi registrado (fica pendente até reiniciar)
        $pending = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" `
                       -Name "Domain" -ErrorAction SilentlyContinue
        if ($pending.Domain -like "*$($domainName.Split('.')[0])*") {
            Write-Status "Ingresso no domínio '$domainName' confirmado! Reinicialize para aplicar." Green
        } else {
            Write-Status "Comando concluído sem erros. Reinicialize para verificar o ingresso." Green
        }

        if (Confirm-Action "Reiniciar agora?") { Restart-Computer -Force }
    }
    catch {
        Write-Err "Falha ao ingressar no domínio '$domainName':"
        Write-Err "$_"
        Write-Host ""
        Write-Info "Causas comuns:"
        Write-Host "   - DNS não aponta para o DC do domínio" -ForegroundColor DarkYellow
        Write-Host "   - Credenciais incorretas ou sem permissão" -ForegroundColor DarkYellow
        Write-Host "   - Computador não alcança o DC na rede"    -ForegroundColor DarkYellow
        Write-Host "   - Nome do computador já existe no AD"     -ForegroundColor DarkYellow
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
$adDomain = ""
$adBaseDN = ""
$adDC = ""
$adUser = ""
$adConnected = $false

function Connect-ADSession {
    if ($script:adConnected -and $script:adSession) { return $true }

    Write-Host ""
    Write-Host "  ==============================================" -ForegroundColor Cyan
    Write-Host "     AUTENTICAÇÃO NO CONTROLADOR DE DOMÍNIO    " -ForegroundColor Yellow
    Write-Host "  ==============================================" -ForegroundColor Cyan
    Write-Host ""

    $script:adDomain = [string](Read-Host "  Domínio (ex: empresa.local)")

    # Valida IP do DC
    do {
        $inputDC = [string](Read-Host "  IP do Controlador de Domínio (ex: 10.0.0.1)")
        $validIP = $false
        if ($inputDC -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $badOctet = $inputDC -split '\.' | Where-Object { [int]$_ -gt 255 }
            $validIP  = ($null -eq $badOctet -or @($badOctet).Count -eq 0)
        }
        if (-not $validIP) { Write-Err "Endereço IP inválido. Use o formato: 10.0.0.1" }
    } while (-not $validIP)
    $script:adDC = $inputDC

    $script:adUser   = [string](Read-Host "  Usuário administrador (ex: Administrator)")

    # Deriva BaseDN automaticamente (ex: empresa.local -> DC=empresa,DC=local)
    $script:adBaseDN = "DC=" + ($script:adDomain.Replace('.', ',DC='))

    $adminPass = Read-Host "  Senha" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential(
        "$($script:adDomain)\$($script:adUser)", $adminPass)

    try {
        # Habilita e configura WinRM de uma vez (inicia serviço + firewall + TrustedHosts)
        Write-Info "Configurando WinRM..."
        Enable-PSRemoting -Force -SkipNetworkProfileCheck -ErrorAction SilentlyContinue 2>&1 | Out-Null

        # Usa variável local tipada como string para evitar Object[] no WSMan
        [string]$dcStr = $script:adDC
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $dcStr -Force -ErrorAction SilentlyContinue

        Write-Info "Conectando ao DC $dcStr..."
        # Negotiate usa NTLM quando Kerberos não funciona (conexão via IP)
        $script:adSession = New-PSSession `
            -ComputerName $dcStr `
            -Credential   $cred  `
            -Authentication Negotiate `
            -ErrorAction Stop
        Invoke-Command -Session $script:adSession -ScriptBlock { Import-Module ActiveDirectory } -ErrorAction Stop
        $script:adConnected = $true
        Write-Status "Conectado ao DC $dcStr ($($script:adDomain)) com sucesso." Green
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

function Select-ADOUDynamic {
    param([string]$prompt = "  Selecione a OU de destino")
    try {
        $allOUs = Invoke-Command -Session $script:adSession -ScriptBlock {
            Get-ADOrganizationalUnit -Filter * -Properties CanonicalName |
                Sort-Object CanonicalName | Select-Object Name, DistinguishedName, CanonicalName
        }
        if (-not $allOUs -or @($allOUs).Count -eq 0) { Write-Err "Nenhuma OU encontrada no domínio."; return $null }

        Write-Host ""
        $i = 1; $ouMap = @{}
        foreach ($ou in @($allOUs)) {
            Write-Host "  [$i] $($ou.CanonicalName)" -ForegroundColor Gray
            $ouMap[$i] = [string]$ou.DistinguishedName
            $i++
        }
        Write-Host "  [0] Cancelar"
        Write-Host ""
        $sel = Read-Host $prompt
        if ($sel -eq '0' -or [string]::IsNullOrWhiteSpace($sel)) { return $null }
        $selInt = [int]$sel
        if (-not $ouMap.ContainsKey($selInt)) { Write-Err "Opção inválida."; return $null }
        return $ouMap[$selInt]
    }
    catch { Write-Err "Erro ao buscar OUs: $_"; return $null }
}

function AD-MoveObject {
    $obj = Read-Host "  Nome do objeto (usuário ou computador)"
    Write-Info "Buscando OUs disponíveis..."
    $targetOU = Select-ADOUDynamic "  Selecione a OU de destino"
    if (-not $targetOU) { return }
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($o, $t)
            $found = @(Get-ADObject -Filter { Name -eq $o } -ErrorAction Stop)
            if ($found.Count -eq 0) { throw "Objeto '$o' não encontrado." }
            if ($found.Count -gt 1)  { throw "Múltiplos objetos com o nome '$o'. Seja mais específico." }
            $found[0] | Move-ADObject -TargetPath $t -ErrorAction Stop
        } -ArgumentList $obj, $targetOU
        Write-Status "Objeto '$obj' movido com sucesso." Green
    }
    catch { Write-Err "Erro: $_" }
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
    Write-Host ""
    Write-Host "  -- Listar Usuários --" -ForegroundColor Cyan
    Write-Host "  [1] Todos os usuários do domínio"
    Write-Host "  [2] Buscar por nome ou login"
    Write-Host "  [3] Filtrar por OU"
    Write-Host ""
    $opt = Read-Host "  Selecione"

    # ScriptBlock reutilizável: busca usuários + grupos em uma única chamada remota
    $fetchBlock = {
        param($filter, $searchBase)
        $props = 'DisplayName','SamAccountName','EmailAddress','Enabled','LastLogonDate','MemberOf'
        $params = @{ Filter = $filter; Properties = $props }
        if ($searchBase) { $params['SearchBase'] = $searchBase }
        Get-ADUser @params | Sort-Object DisplayName | ForEach-Object {
            $groups = @($_.MemberOf) |
                ForEach-Object { ($_ -split ',')[0] -replace '^CN=','' } |
                Sort-Object
            [PSCustomObject]@{
                DisplayName    = $_.DisplayName
                SamAccountName = $_.SamAccountName
                EmailAddress   = $_.EmailAddress
                Enabled        = $_.Enabled
                LastLogonDate  = $_.LastLogonDate
                Groups         = $groups
            }
        }
    }

    try {
        $users = $null

        if ($opt -eq '1') {
            Write-Info "Buscando todos os usuários do domínio..."
            $users = Invoke-Command -Session $script:adSession -ScriptBlock $fetchBlock `
                -ArgumentList '*', $null
        }
        elseif ($opt -eq '2') {
            $termo = Read-Host "  Nome ou login (parcial, ex: joao)"
            Write-Info "Buscando '$termo'..."
            $users = Invoke-Command -Session $script:adSession -ScriptBlock {
                param($t)
                $props = 'DisplayName','SamAccountName','EmailAddress','Enabled','LastLogonDate','MemberOf'
                Get-ADUser -Filter { (Name -like $t) -or (SamAccountName -like $t) } `
                    -Properties $props | Sort-Object DisplayName | ForEach-Object {
                    $groups = @($_.MemberOf) |
                        ForEach-Object { ($_ -split ',')[0] -replace '^CN=','' } |
                        Sort-Object
                    [PSCustomObject]@{
                        DisplayName    = $_.DisplayName
                        SamAccountName = $_.SamAccountName
                        EmailAddress   = $_.EmailAddress
                        Enabled        = $_.Enabled
                        LastLogonDate  = $_.LastLogonDate
                        Groups         = $groups
                    }
                }
            } -ArgumentList "*$termo*"
        }
        elseif ($opt -eq '3') {
            Write-Info "Buscando OUs disponíveis..."
            $ouPath = Select-ADOUDynamic "  Selecione a OU"
            if (-not $ouPath) { return }
            Write-Info "Buscando usuários em '$ouPath'..."
            $users = Invoke-Command -Session $script:adSession -ScriptBlock $fetchBlock `
                -ArgumentList '*', $ouPath
        }
        else { Write-Err "Opção inválida."; return }

        $list = @($users)
        if ($list.Count -eq 0) { Write-Info "Nenhum usuário encontrado."; return }

        Write-Host ""
        Write-Host "  $('=' * 100)" -ForegroundColor DarkCyan

        foreach ($u in $list) {
            $status = if ($u.Enabled) { "Ativo  " } else { "Inativo" }
            $color  = if ($u.Enabled) { "White" }   else { "DarkGray" }
            $logon  = if ($u.LastLogonDate) { $u.LastLogonDate.ToString("dd/MM/yyyy") } else { "Nunca" }
            $nome   = if ($u.DisplayName)   { $u.DisplayName } else { $u.SamAccountName }

            Write-Host ("  {0,-30} {1,-22} {2,-34} {3}  Logon: {4}" -f `
                $nome, $u.SamAccountName, $u.EmailAddress, $status, $logon) -ForegroundColor $color

            $grps = @($u.Groups)
            if ($grps.Count -gt 0) {
                Write-Host ("    └─ Grupos ({0}): {1}" -f $grps.Count, ($grps -join ' | ')) `
                    -ForegroundColor DarkYellow
            } else {
                Write-Host "    └─ Grupos: nenhum" -ForegroundColor DarkGray
            }
        }

        Write-Host "  $('=' * 100)" -ForegroundColor DarkCyan
        Write-Host "  Total: $($list.Count) usuário(s)" -ForegroundColor Yellow
    }
    catch { Write-Err "Erro ao listar usuários: $_" }
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
    Write-Host ""
    Write-Host "  [1] Delta  - sincroniza apenas alterações recentes (recomendado)"
    Write-Host "  [2] Full   - sincronização completa"
    Write-Host ""
    $tipo = Read-Host "  Selecione"
    $policy = if ($tipo -eq '2') { 'Initial' } else { 'Delta' }
    Write-Info "Iniciando sincronização $policy..."
    try {
        Invoke-Command -Session $script:adSession -ScriptBlock {
            param($p)
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType $p -ErrorAction Stop
        } -ArgumentList $policy -ErrorAction Stop
        Write-Status "Sincronização $policy iniciada com sucesso no DC." Green
    }
    catch { Write-Err "Erro: $_`n  Verifique se o Azure AD Connect está instalado no DC." }
}

function AD-ExportUsers {
    Write-Host ""
    Write-Host "  [1] Todos os usuários do domínio"
    Write-Host "  [2] Filtrar por OU"
    Write-Host ""
    $opt = Read-Host "  Selecione"

    $ouPath = $null
    if ($opt -eq '2') {
        Write-Info "Buscando OUs disponíveis..."
        $ouPath = Select-ADOUDynamic "  Selecione a OU"
        if (-not $ouPath) { return }
    }

    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $dest = "$env:USERPROFILE\Desktop\usuarios_$ts.csv"
    $file = Read-Host "  Arquivo de saída (Enter = Desktop\usuarios_$ts.csv)"
    if ([string]::IsNullOrWhiteSpace($file)) { $file = $dest }

    try {
        $users = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($p)
            $filter = if ($p) {
                Get-ADUser -Filter * -SearchBase $p -Properties DisplayName, SamAccountName, EmailAddress, Title, Department, Enabled, LastLogonDate
            } else {
                Get-ADUser -Filter * -Properties DisplayName, SamAccountName, EmailAddress, Title, Department, Enabled, LastLogonDate
            }
            $filter | Select-Object DisplayName, SamAccountName, EmailAddress, Title, Department, Enabled, LastLogonDate |
                Sort-Object DisplayName
        } -ArgumentList $ouPath

        @($users) | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        Write-Status "Exportado $(@($users).Count) usuário(s) para: $file" Green
    }
    catch { Write-Err "Erro: $_" }
}

function AD-ExportComputers {
    Write-Host ""
    Write-Host "  [1] Todos os computadores do domínio"
    Write-Host "  [2] Filtrar por OU"
    Write-Host ""
    $opt = Read-Host "  Selecione"

    $ouPath = $null
    if ($opt -eq '2') {
        Write-Info "Buscando OUs disponíveis..."
        $ouPath = Select-ADOUDynamic "  Selecione a OU"
        if (-not $ouPath) { return }
    }

    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $dest = "$env:USERPROFILE\Desktop\computadores_$ts.csv"
    $file = Read-Host "  Arquivo de saída (Enter = Desktop\computadores_$ts.csv)"
    if ([string]::IsNullOrWhiteSpace($file)) { $file = $dest }

    try {
        $computers = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($p)
            $filter = if ($p) {
                Get-ADComputer -Filter * -SearchBase $p -Properties Name, DNSHostName, Enabled, LastLogonDate, OperatingSystem
            } else {
                Get-ADComputer -Filter * -Properties Name, DNSHostName, Enabled, LastLogonDate, OperatingSystem
            }
            $filter | Select-Object Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate |
                Sort-Object Name
        } -ArgumentList $ouPath

        @($computers) | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        Write-Status "Exportado $(@($computers).Count) computador(es) para: $file" Green
    }
    catch { Write-Err "Erro: $_" }
}

function AD-ExportGroups {
    Write-Host ""
    Write-Host "  [1] Todos os grupos do domínio"
    Write-Host "  [2] Filtrar por OU"
    Write-Host ""
    $opt = Read-Host "  Selecione"

    $ouPath = $null
    if ($opt -eq '2') {
        Write-Info "Buscando OUs disponíveis..."
        $ouPath = Select-ADOUDynamic "  Selecione a OU"
        if (-not $ouPath) { return }
    }

    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $dest = "$env:USERPROFILE\Desktop\grupos_$ts.csv"
    $file = Read-Host "  Arquivo de saída (Enter = Desktop\grupos_$ts.csv)"
    if ([string]::IsNullOrWhiteSpace($file)) { $file = $dest }

    try {
        $groups = Invoke-Command -Session $script:adSession -ScriptBlock {
            param($p)
            $filter = if ($p) {
                Get-ADGroup -Filter * -SearchBase $p -Properties Name, SamAccountName, GroupCategory, GroupScope, Description
            } else {
                Get-ADGroup -Filter * -Properties Name, SamAccountName, GroupCategory, GroupScope, Description
            }
            $filter | Select-Object Name, SamAccountName, GroupCategory, GroupScope, Description |
                Sort-Object Name
        } -ArgumentList $ouPath

        @($groups) | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        Write-Status "Exportado $(@($groups).Count) grupo(s) para: $file" Green
    }
    catch { Write-Err "Erro: $_" }
}

function Manage-AD {
    if (-not (Connect-ADSession)) { return }

    do {
        Write-Host ""
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host "         GERENCIAMENTO DO ACTIVE DIRECTORY      " -ForegroundColor Yellow
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host "  Domínio: $($script:adDomain)  |  DC: $($script:adDC)  |  Usuário: $($script:adUser)" -ForegroundColor Green
        Write-Host "  ----------------------------------------------"
        Write-Host "  -- Usuários --" -ForegroundColor DarkCyan
        Write-Host "   [1]  Criar usuário"
        Write-Host "   [2]  Inativar usuário"
        Write-Host "   [3]  Reativar usuário"
        Write-Host "   [4]  Deletar usuário"
        Write-Host "   [5]  Resetar senha"
        Write-Host "   [6]  Desbloquear usuário"
        Write-Host "   [7]  Listar usuários"
        Write-Host "  -- Computadores --" -ForegroundColor DarkCyan
        Write-Host "   [8]  Adicionar computador"
        Write-Host "   [9]  Remover computador"
        Write-Host "  [10]  Listar computadores"
        Write-Host "  -- Grupos --" -ForegroundColor DarkCyan
        Write-Host "  [11]  Adicionar usuário a grupo"
        Write-Host "  [12]  Remover usuário de grupo"
        Write-Host "  [13]  Ver membros de grupo"
        Write-Host "  -- Outros --" -ForegroundColor DarkCyan
        Write-Host "  [14]  Mover objeto para outra OU"
        Write-Host "  [15]  Sincronizar AD"
        Write-Host "  [16]  Exportar relatório de usuários"
        Write-Host "  [17]  Exportar relatório de computadores"
        Write-Host "  [18]  Exportar relatório de grupos"
        Write-Host "   [0]  Voltar ao menu principal" -ForegroundColor Red
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host ""

        $adChoice = Read-Host "  Selecione"
        switch ($adChoice) {
            '1'  { AD-CreateUser }
            '2'  { AD-DisableUser }
            '3'  { AD-EnableUser }
            '4'  { AD-DeleteUser }
            '5'  { AD-ResetPassword }
            '6'  { AD-UnlockUser }
            '7'  { AD-ListUsers }
            '8'  { AD-AddComputer }
            '9'  { AD-RemoveComputer }
            '10' { AD-ListComputers }
            '11' { AD-AddUserToGroup }
            '12' { AD-RemoveUserFromGroup }
            '13' { AD-GetGroupMembers }
            '14' { AD-MoveObject }
            '15' { AD-SyncAD }
            '16' { AD-ExportUsers }
            '17' { AD-ExportComputers }
            '18' { AD-ExportGroups }
            '0'  { Disconnect-ADSession; break }
            default { Write-Err "Opção inválida." }
        }

        if ($adChoice -ne '0') {
            Write-Host ""
            Read-Host "  Pressione Enter para continuar"
        }

    } while ($adChoice -ne '0')
}

# -------------------------------------------
#  8. SHAREPOINT ACCESS MANAGEMENT
# -------------------------------------------

$spTenantUrl  = ""
$spConnected  = $false
# PnP Management Shell — app público multi-tenant da Microsoft/PnP.
# IMPORTANTE: requer consentimento admin no tenant antes do primeiro uso (opção [2]).
# Alternativa recomendada: registre um app próprio no Azure e use a opção [3].
$spClientId_Default = "31359c7f-bd7e-475c-86db-fdb8c937548e"
$spClientId         = $spClientId_Default
$spConfigFile       = "$env:APPDATA\WinUtil\sp_config.json"

function SP-LoadConfig {
    if (Test-Path $spConfigFile) {
        try {
            $cfg = Get-Content $spConfigFile -Raw | ConvertFrom-Json
            if ($cfg.ClientId -and $cfg.ClientId -ne "") {
                $script:spClientId = $cfg.ClientId
                return $cfg
            }
        }
        catch {}
    }
    return $null
}

function SP-SaveConfig {
    param([string]$clientId)
    $dir = Split-Path $spConfigFile
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    @{ ClientId = $clientId } | ConvertTo-Json | Set-Content $spConfigFile -Encoding UTF8
}

function Ensure-PnPModule {
    if (Get-Module -ListAvailable -Name 'PnP.PowerShell') {
        Import-Module PnP.PowerShell -ErrorAction SilentlyContinue
        return $true
    }
    Write-Info "PnP.PowerShell não encontrado. Instalando do PSGallery..."
    try {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module PnP.PowerShell -ErrorAction Stop
        Write-Status "PnP.PowerShell instalado com sucesso." Green
        return $true
    }
    catch {
        Write-Err "Falha ao instalar PnP.PowerShell: $_"
        Write-Info "Instalação manual: Install-Module PnP.PowerShell -Scope CurrentUser"
        return $false
    }
}

function SP-NormalizeAdminUrl {
    param([string]$url)
    $url = $url.TrimEnd('/', '\', ':').Trim()
    if ($url -match '^(https?://[^/]+)') { $url = $Matches[1] }
    # Converte URL raiz para URL de admin  (https://tenant.sharepoint.com → https://tenant-admin.sharepoint.com)
    if ($url -match '^https://([^-][^.]+)\.sharepoint\.com$') {
        $url = "https://$($Matches[1])-admin.sharepoint.com"
    }
    return $url
}

function SP-ShowAuthHelp {
    Write-Host ""
    Write-Host "  +---------------------------------------------------------------+" -ForegroundColor DarkYellow
    Write-Host "  |  ERRO DE AUTENTICAÇÃO — Como resolver:                        |" -ForegroundColor Yellow
    Write-Host "  +---------------------------------------------------------------+" -ForegroundColor DarkYellow
    Write-Host "  |                                                               |" -ForegroundColor DarkYellow
    Write-Host "  |  OPÇÃO A — Consentimento do app PnP (mais rápido):            |" -ForegroundColor Cyan
    Write-Host "  |    1. Volte ao menu e escolha opção [2]                       |" -ForegroundColor White
    Write-Host "  |    2. Entre com uma conta Global Admin                        |" -ForegroundColor White
    Write-Host "  |    3. Aceite as permissões solicitadas                        |" -ForegroundColor White
    Write-Host "  |    4. Feito! Tente conectar novamente com opção [1]           |" -ForegroundColor White
    Write-Host "  |                                                               |" -ForegroundColor DarkYellow
    Write-Host "  |  OPÇÃO B — App próprio no Azure (recomendado, mais seguro):   |" -ForegroundColor Cyan
    Write-Host "  |    1. Acesse portal.azure.com → Entra ID → App registrations |" -ForegroundColor White
    Write-Host "  |    2. New registration → nome: WinUtil SP → Public client     |" -ForegroundColor White
    Write-Host "  |    3. API permissions → SharePoint → Sites.FullControl.All   |" -ForegroundColor White
    Write-Host "  |       (Delegated) + Grant admin consent                       |" -ForegroundColor White
    Write-Host "  |    4. Copie o Application (client) ID                        |" -ForegroundColor White
    Write-Host "  |    5. Volte ao menu e use a opção [3] para salvar o ID        |" -ForegroundColor White
    Write-Host "  +---------------------------------------------------------------+" -ForegroundColor DarkYellow
    Write-Host ""
}

function SP-Connect {
    if ($script:spConnected) { return $true }

    # Carrega configuração salva (ClientId customizado, se houver)
    SP-LoadConfig | Out-Null

    Write-Host ""
    Write-Info "Informe qualquer URL do SharePoint do seu tenant."
    $inputUrl = Read-Host "  SharePoint URL (ex: https://cinnovation.sharepoint.com)"
    if ([string]::IsNullOrWhiteSpace($inputUrl)) { Write-Err "URL não pode ser vazia."; return $false }
    $script:spTenantUrl = SP-NormalizeAdminUrl $inputUrl

    $clientLabel = if ($script:spClientId -eq $script:spClientId_Default) { "PnP Management Shell (padrão)" } else { "App customizado: $($script:spClientId)" }
    Write-Info "Admin URL  : $($script:spTenantUrl)"
    Write-Info "Client App : $clientLabel"
    Write-Info "Uma janela do browser será aberta para login Microsoft..."

    try {
        Connect-PnPOnline -Url $script:spTenantUrl -Interactive -ClientId $script:spClientId -ErrorAction Stop
        $script:spConnected = $true
        Write-Status "Conectado ao SharePoint Admin com sucesso." Green
        return $true
    }
    catch {
        $errMsg = $_.ToString()

        # Erro AADSTS700016 = app não consentido no tenant
        if ($errMsg -like "*AADSTS700016*") {
            Write-Host ""
            Write-Err "O app '$($script:spClientId)' não está autorizado no seu tenant."
            SP-ShowAuthHelp
        }
        # Erro AADSTS50011 = redirect URI inválida
        elseif ($errMsg -like "*AADSTS50011*") {
            Write-Err "Redirect URI inválida. Verifique as configurações do app no Azure."
            Write-Info "Para app customizado: adicione 'http://localhost' em Authentication → Redirect URIs."
        }
        # Erro de acesso negado / permissão
        elseif ($errMsg -like "*Forbidden*" -or $errMsg -like "*Access*denied*" -or $errMsg -like "*403*") {
            Write-Err "Acesso negado. A conta não tem permissão de SharePoint Admin."
            Write-Info "Use uma conta com papel de SharePoint Administrator ou Global Admin."
        }
        else {
            Write-Err "Falha na conexão: $errMsg"
            SP-ShowAuthHelp
        }
        return $false
    }
}

function SP-MirrorPermissions {
    Write-Section "SHAREPOINT — ESPELHAR PERMISSÕES"
    Write-Host ""
    Write-Host "  Copia todas as permissões SharePoint de um colaborador existente" -ForegroundColor Gray
    Write-Host "  para um novo colaborador (onboarding)." -ForegroundColor Gray
    Write-Host ""

    $sourceUPN = Read-Host "  UPN do colaborador ORIGEM  (copiar DE, ex: joao.silva@empresa.com)"
    $targetUPN = Read-Host "  UPN do colaborador DESTINO (copiar PARA, ex: maria.souza@empresa.com)"

    if ([string]::IsNullOrWhiteSpace($sourceUPN) -or [string]::IsNullOrWhiteSpace($targetUPN)) {
        Write-Err "UPN não pode ser vazio."
        return
    }

    Write-Host ""
    Write-Host "  Modo:" -ForegroundColor Cyan
    Write-Host "  [1] Aplicar permissões (altera o ambiente)"
    Write-Host "  [2] Apenas visualizar (sem alterações)"
    $modeChoice  = Read-Host "  Selecione"
    $previewOnly = ($modeChoice -eq '2')
    if ($previewOnly) { Write-Info "Modo visualização — nenhuma alteração será feita." }

    Write-Host ""
    Write-Info "Buscando todos os sites do tenant (pode levar alguns instantes)..."

    try {
        $allSites = Get-PnPTenantSite -ErrorAction Stop | Where-Object {
            $_.Template -ne 'RedirectSite#0' -and
            $_.Template -ne 'SRCHCEN#0'      -and
            $_.Url -notlike '*-my.sharepoint.com*'
        }
    }
    catch {
        $errMsg = $_.ToString()
        if ($errMsg -like "*AADSTS700016*") {
            Write-Err "App não autorizado no tenant."
            SP-ShowAuthHelp
        }
        else {
            Write-Err "Falha ao buscar sites: $errMsg"
        }
        return
    }

    $totalSites = @($allSites).Count
    Write-Info "Encontrado(s) $totalSites site(s). Verificando permissões de '$sourceUPN'..."

    if ($totalSites -eq 0) {
        Write-Err "Nenhum site encontrado. Verifique se a conta tem permissão de SharePoint Admin."
        return
    }

    if ($totalSites -gt 50 -and -not $previewOnly) {
        if (-not (Confirm-Action "$totalSites sites encontrados — pode demorar vários minutos. Continuar?")) { return }
    }

    Write-Host ""

    $results   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $siteIdx   = 0
    $padWidth  = ([string]$totalSites).Length
    $errCount  = 0

    foreach ($site in $allSites) {
        $siteIdx++
        Write-Host ("  [{0,$padWidth}/{1}] {2}" -f $siteIdx, $totalSites, $site.Url) -ForegroundColor DarkGray -NoNewline

        try {
            # MSAL reutiliza token em cache — sem popup adicional após o primeiro login
            Connect-PnPOnline -Url $site.Url -Interactive -ClientId $script:spClientId -ErrorAction Stop 2>$null
        }
        catch {
            Write-Host " [falha de conexão]" -ForegroundColor DarkRed
            $errCount++
            continue
        }

        $copied = [System.Collections.Generic.List[string]]::new()

        # Site Collection Admins
        try {
            $admins  = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
            $isAdmin = @($admins) | Where-Object { $_.Email -eq $sourceUPN -or $_.LoginName -like "*$sourceUPN*" }
            if ($isAdmin) {
                if (-not $previewOnly) { Add-PnPSiteCollectionAdmin -Owners $targetUPN -ErrorAction SilentlyContinue }
                $copied.Add("Site Collection Admin")
            }
        }
        catch {}

        # Membros de grupos SharePoint
        try {
            $groups = Get-PnPGroup -ErrorAction SilentlyContinue
            foreach ($group in @($groups)) {
                try {
                    $members = Get-PnPGroupMember -Identity $group.Title -ErrorAction SilentlyContinue
                    $inGroup = @($members) | Where-Object { $_.Email -eq $sourceUPN -or $_.LoginName -like "*$sourceUPN*" }
                    if ($inGroup) {
                        if (-not $previewOnly) { Add-PnPGroupMember -LoginName $targetUPN -Group $group.Title -ErrorAction SilentlyContinue }
                        $copied.Add("Grupo: $($group.Title)")
                    }
                }
                catch {}
            }
        }
        catch {}

        if ($copied.Count -gt 0) {
            $verb = if ($previewOnly) { "encontrado(s)" } else { "copiado(s)" }
            Write-Host (" → $($copied.Count) permissão(ões) $verb") -ForegroundColor Green
            $results.Add([PSCustomObject]@{ Site = $site.Url; Permissoes = $copied -join '; ' })
        }
        else {
            Write-Host ""
        }
    }

    Write-Host ""

    if ($errCount -gt 0) {
        Write-Info "$errCount site(s) com falha de conexão foram ignorados."
    }

    if ($results.Count -gt 0) {
        $verb = if ($previewOnly) { "encontradas" } else { "copiadas" }
        Write-Status "Concluído! Permissões $verb em $($results.Count) site(s):" Green
        Write-Host ""
        foreach ($r in $results) {
            Write-Host "  $($r.Site)" -ForegroundColor Cyan
            Write-Host "    +-- $($r.Permissoes)" -ForegroundColor Gray
        }
        Write-Host ""
        if (Confirm-Action "Exportar relatório para o Desktop?") {
            $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
            $file = "$env:USERPROFILE\Desktop\sp_mirror_$ts.csv"
            $results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
            Write-Status "Relatório salvo em: $file" Green
        }
    }
    else {
        Write-Info "Nenhuma permissão encontrada para '$sourceUPN' nos $totalSites site(s) verificados."
        Write-Info "Certifique-se de que o UPN está correto (formato: usuario@dominio.com)."
    }
}

function SP-GrantConsent {
    Write-Section "CONSENTIMENTO DO APP PnP MANAGEMENT SHELL"
    Write-Host ""
    Write-Host "  Este processo registra o app PnP Management Shell no seu tenant Entra ID." -ForegroundColor Gray
    Write-Host "  É necessário apenas uma vez. Requer conta de Global Administrator." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Após o consentimento, qualquer usuário do tenant poderá usar a opção [1]." -ForegroundColor DarkYellow
    Write-Host ""
    if (-not (Confirm-Action "Abrir browser para consentimento como Global Admin?")) { return }
    try {
        Register-PnPManagementShellAccess -ErrorAction Stop
        Write-Status "Consentimento concedido! Use a opção [1] para conectar." Green
    }
    catch {
        $errMsg = $_.ToString()
        Write-Err "Falha: $errMsg"
        if ($errMsg -like "*Register-PnPManagementShellAccess*" -or $errMsg -like "*não reconhecido*") {
            Write-Info "Versão do PnP.PowerShell pode estar desatualizada."
            Write-Info "Atualize: Update-Module PnP.PowerShell"
        }
        Write-Info "Alternativa manual: acesse https://aka.ms/pnp-auth-consent para consentir pelo browser."
    }
}

function SP-ConfigureCustomApp {
    Write-Section "CONFIGURAR APP AZURE CUSTOMIZADO"
    Write-Host ""
    Write-Host "  Use esta opção se você registrou um app próprio no Azure portal." -ForegroundColor Gray
    Write-Host "  Isso é mais seguro e não depende de consentimento global do PnP." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Como registrar o app (se ainda não fez):" -ForegroundColor Cyan
    Write-Host "  1. portal.azure.com → Entra ID → App registrations → New registration"
    Write-Host "  2. Nome: WinUtil SharePoint | Tipo: Public client (mobile and desktop)"
    Write-Host "  3. API permissions → SharePoint → Delegated → Sites.FullControl.All"
    Write-Host "     + Microsoft Graph → Delegated → User.Read"
    Write-Host "  4. Grant admin consent → Copie o Application (client) ID"
    Write-Host ""

    $currentId = if ($script:spClientId -ne $script:spClientId_Default) { $script:spClientId } else { "(nenhum configurado)" }
    Write-Info "Client ID atual: $currentId"
    Write-Host ""

    $newId = Read-Host "  Cole o Application (client) ID do seu app (Enter para cancelar)"
    if ([string]::IsNullOrWhiteSpace($newId)) { Write-Info "Cancelado."; return }

    # Valida formato GUID
    if ($newId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        Write-Err "Formato inválido. O Client ID deve ser um GUID (ex: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)."
        return
    }

    $script:spClientId = $newId
    SP-SaveConfig -clientId $newId
    $script:spConnected = $false  # força reconexão com novo app
    Write-Status "App customizado configurado! Client ID: $newId" Green
    Write-Info "A próxima conexão usará este app. A configuração foi salva em: $spConfigFile"
}

function SP-ResetConfig {
    $script:spClientId  = $script:spClientId_Default
    $script:spConnected = $false
    if (Test-Path $spConfigFile) { Remove-Item $spConfigFile -Force }
    Write-Status "Configuração resetada para PnP Management Shell (padrão)." Green
}

function Manage-SharePoint {
    if (-not (Ensure-PnPModule)) { return }

    # Carrega ClientId salvo, se existir
    SP-LoadConfig | Out-Null

    do {
        Write-Host ""
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host "         SHAREPOINT ACCESS MANAGEMENT           " -ForegroundColor Yellow
        Write-Host "  ==============================================" -ForegroundColor Cyan
        if ($script:spConnected) {
            Write-Host "  Status : Conectado" -ForegroundColor Green
            Write-Host "  Tenant : $($script:spTenantUrl)" -ForegroundColor Green
        } else {
            Write-Host "  Status : Não conectado" -ForegroundColor Yellow
        }
        $appLabel = if ($script:spClientId -eq $script:spClientId_Default) { "PnP Management Shell (padrão)" } else { "App customizado" }
        Write-Host "  App    : $appLabel" -ForegroundColor Gray
        Write-Host "  ----------------------------------------------"
        Write-Host "  [1]  Espelhar permissões (onboarding de novo colaborador)" -ForegroundColor Green
        Write-Host "  [2]  Consentimento inicial — PnP Management Shell (Global Admin)" -ForegroundColor DarkYellow
        Write-Host "  [3]  Configurar app Azure próprio (recomendado)" -ForegroundColor Cyan
        Write-Host "  [4]  Resetar para configuração padrão" -ForegroundColor DarkGray
        Write-Host "  [0]  Voltar ao menu principal" -ForegroundColor Red
        Write-Host "  ==============================================" -ForegroundColor Cyan
        Write-Host ""

        $spChoice = Read-Host "  Selecione"
        switch ($spChoice) {
            '1'  { if (SP-Connect) { SP-MirrorPermissions } }
            '2'  { SP-GrantConsent }
            '3'  { SP-ConfigureCustomApp }
            '4'  { SP-ResetConfig }
            '0'  { $script:spConnected = $false; Disconnect-PnPOnline -ErrorAction SilentlyContinue; break }
            default { Write-Err "Opção inválida." }
        }

        if ($spChoice -ne '0') {
            Write-Host ""
            Read-Host "  Pressione Enter para continuar"
        }

    } while ($spChoice -ne '0')
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
    Write-Host "  |  [8]  SharePoint Management        |" -ForegroundColor Cyan
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
        '8' { Manage-SharePoint }
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
