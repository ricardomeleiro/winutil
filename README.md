# WinUtil - Windows Utility Tool

> One command to set up, configure and repair any Windows machine.

```powershell
irm https://win.c-innovation.com.br/win.ps1 | iex
```

> **Run PowerShell as Administrator before executing.**

---

## What it does

WinUtil is an interactive PowerShell script that runs entirely in memory — no installer, no executable, no bloat. It provides a menu-driven interface with 6 modules:

| # | Module | Description |
|---|--------|-------------|
| 1 | **Default Softwares** | Installs the standard company software kit in one shot |
| 2 | **App Installer** | Pick and install any app from a curated list via winget |
| 3 | **Windows Tweaks** | Privacy, performance and UI/UX registry tweaks |
| 4 | **System Diagnostics** | Full hardware/software snapshot and top process view |
| 5 | **Fix and Repair** | SFC, DISM, DNS flush, Windows Update reset and more |
| 6 | **Domain Management** | Join/leave Onshore and Offshore domains |

---

## Module details

### [1] Default Softwares
Installs the following apps silently with no interaction required:

- Microsoft 365 (Office)
- Microsoft Teams
- Google Chrome
- Lightshot
- WinRAR
- WhatsApp
- Power BI Desktop

### [2] App Installer
Choose from 25+ apps organized by category:
- Browsers (Chrome, Firefox, Brave, Opera GX)
- Dev Tools (VS Code, Git, Node.js, Python, Docker, Windows Terminal)
- Communication (Discord, Slack, Teams, Zoom)
- Utilities (7-Zip, VLC, Everything, Notepad++, PowerToys)
- Security (Malwarebytes, Bitwarden, WireGuard)

### [3] Windows Tweaks
- **Privacy** — disables telemetry, advertising ID, Cortana, activity tracking
- **Performance** — sets High Performance power plan, disables SysMain, removes startup delay
- **UI/UX** — enables dark mode, shows file extensions and hidden files, cleans up taskbar

### [4] System Diagnostics
Displays a full system overview including OS build, CPU, RAM, GPU, disk usage, network adapter, IP address, hostname, uptime, top 5 processes by CPU, and Windows Defender status.

### [5] Fix and Repair
- SFC + DISM system file repair
- Windows Update cache reset
- DNS flush and full network stack reset
- Temp file and disk cleanup
- Windows Store cache reset
- Option to run all repairs at once

### [6] Domain Management
- **Join ONSHORE** — `chouest-br.local`
- **Join OFFSHORE** — `rov.local`
- **Join custom domain**
- **Leave domain** and join a workgroup
- **View** current domain, DC info and local users

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- **Run as Administrator**
- Internet connection (for app installs)
- `winget` — App Installer from the Microsoft Store (the script will prompt you to install it if missing)

---

## How it works

```
User runs: irm https://win.c-innovation.com.br/win.ps1 | iex
                          |
                          v
             Apache reverse proxy (EC2)
                          |
                          v
     raw.githubusercontent.com/YOUR_USER/winutil/main/win.ps1
                          |
                          v
         Script runs entirely in memory on the local machine
```

The web server proxies requests to this GitHub repository, so the command always executes the **latest committed version** of the script — no server changes needed after a GitHub commit.

---

## Updating the script

1. Edit `win.ps1` directly on GitHub (or clone, edit, push)
2. Commit the changes
3. Done — the next `irm` command will fetch the new version automatically

---

## Security notice

Always review scripts before running them via `irm | iex`. You can inspect the full source at any time:

- **View raw script:** https://raw.githubusercontent.com/YOUR_USER/winutil/main/win.ps1
- **Browse repo:** https://github.com/YOUR_USER/winutil

This script requires Administrator privileges and modifies Windows registry keys and system services. Read the source before executing on any machine.

---

## License

MIT — free to use, modify and distribute.

---

*Maintained by C-Innovation — https://c-innovation.com.br*
