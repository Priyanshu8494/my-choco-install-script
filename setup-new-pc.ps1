<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
.NOTES
  - Clean UI Edition
  - Work in progress.
#>

# -------------------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------------------
function Show-Header {
    Clear-Host
    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host '     ==/          i     i          \==' -ForegroundColor DarkGray
    Write-Host '   /XX/            |\_/|            \XX\' -ForegroundColor DarkGray
    Write-Host ' /XXXX\            |XXXXX|            /XXXX\' -ForegroundColor DarkGray
    Write-Host '|XXXXXX\         _XXXXXXX         /XXXXXX|' -ForegroundColor DarkGray
    Write-Host 'XXXXXXXXXXxxxxxxxXXXXXXXXXXXxxxxxxxXXXXXXXXXX' -ForegroundColor DarkGray
    Write-Host '|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|' -ForegroundColor DarkGray
    Write-Host 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' -ForegroundColor DarkGray
    Write-Host '|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|' -ForegroundColor DarkGray
    Write-Host ' XXXXXX/^^^^"\XXXXXXXXXXXXXXXXXXXXX/^^^^"\XXXXXX' -ForegroundColor DarkGray
    Write-Host '  |XX|       \XXX/^^\XXXXX/^^\XXX/       |XX|' -ForegroundColor DarkGray
    Write-Host '    \|       \X/    \XXX/    \X/       |/' -ForegroundColor DarkGray
    Write-Host '             !       \X/       !' -ForegroundColor DarkGray
    Write-Host '                     !' -ForegroundColor DarkGray
    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host '     Priyanshu Suryavanshi PC Setup Toolkit' -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================================`n" -ForegroundColor White
}

function Show-Menu {
    param (
        [string]$StatusMessage = "",
        [string]$StatusColor   = "Yellow"
    )
    Show-Header
    Write-Host "📌 Status:" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "   ➤ $StatusMessage" -ForegroundColor $StatusColor
    } else {
        Write-Host "   ➤ Ready to assist with your setup!" -ForegroundColor Green
    }

    Write-Host "`n🧰 Main Menu:" -ForegroundColor Yellow
    Write-Host "   [1] 📦 Install Essential Software" -ForegroundColor White
    Write-Host "   [2] 💼 Install MS Office Suite" -ForegroundColor White
    Write-Host "   [3] 🔑 System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "   [4] 🔄 Update All Installed Software (via Winget)" -ForegroundColor White
    Write-Host "   [5] 🚀 Advanced Toolkit (runs remote script)" -ForegroundColor White
    Write-Host "   [0] ❌ Exit" -ForegroundColor Red
    Write-Host "`n💡 Tip: Use numbers to navigate (e.g., 1 for software install)" -ForegroundColor DarkCyan
    Write-Host "`n============================================================" -ForegroundColor DarkCyan
}

function Show-Loading {
    param([string]$Message = "Processing")
    $dots = ".", "..", "..."
    for ($i = 0; $i -lt $dots.Count; $i++) {
        Write-Host ("`r{0}{1}" -f $Message, $dots[$i]) -NoNewline
        Start-Sleep -Milliseconds 300
    }
    Write-Host "`r$Message... Done!`n" -ForegroundColor Green
}

# -------------------------------------------------------------------------
# Ensure package managers present
# -------------------------------------------------------------------------
function Ensure-PackageManagers {
    # Check winget
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "`n❌ Winget is not installed or not available in PATH!" -ForegroundColor Red
        Write-Host "➡️ Please install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        # Do not automatically exit — allow user to continue for tasks that don't need winget
    } else {
        Write-Host "`n✅ Winget detected." -ForegroundColor Green
    }

    # Check Chocolatey, install if missing
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "🍫 Chocolatey not found. Installing now..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            $chocoInstallScript = 'https://community.chocolatey.org/install.ps1'
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString($chocoInstallScript))
            Write-Host "✅ Chocolatey installation attempted. You may need to re-open shell." -ForegroundColor Green
        } catch {
            Write-Host "❌ Chocolatey install failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "✅ Chocolatey detected." -ForegroundColor Green
    }
}

# -------------------------------------------------------------------------
# Utility: Create Desktop Shortcut
# -------------------------------------------------------------------------
function New-DesktopShortcut {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [Parameter(Mandatory)][string]$ShortcutName,
        [string]$Arguments = ""
    )
    try {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$desktop\$ShortcutName.lnk")
        $Shortcut.TargetPath = $TargetPath
        if ($Arguments) { $Shortcut.Arguments = $Arguments }
        $Shortcut.IconLocation = "$TargetPath,0"
        $Shortcut.Save()
        return $true
    } catch {
        return $false
    }
}

# -------------------------------------------------------------------------
# Software installation
# -------------------------------------------------------------------------
function Install-AnyDeskDirectly {
    $anydeskUrl = "https://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\AnyDesk"
    $exePath = Join-Path $installPath "AnyDesk.exe"

    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan
    try {
        if (-not (Test-Path $installPath)) {
            New-Item -ItemType Directory -Path $installPath -Force | Out-Null
        }
        Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath -UseBasicParsing -ErrorAction Stop

        Write-Host "🚀 Running AnyDesk installer (silent flags may vary)..." -ForegroundColor Yellow
        Start-Process -FilePath $exePath -ArgumentList "--install" -Wait -NoNewWindow -ErrorAction SilentlyContinue

        # Create shortcut (if executable exists)
        if (Test-Path $exePath) {
            if (New-DesktopShortcut -TargetPath $exePath -ShortcutName "AnyDesk") {
                Write-Host "✅ AnyDesk installed and shortcut created." -ForegroundColor Green
            } else {
                Write-Host "✅ AnyDesk downloaded. Shortcut creation failed or skipped." -ForegroundColor Yellow
            }
        } else {
            Write-Host "❌ AnyDesk installer file not found after download." -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ Error installing AnyDesk: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-UltraViewerDirectly {
    $uvUrl = "https://ultraviewer.net/UltraViewer_setup.exe"
    $installPath = "C:\UltraViewer"
    $exePath = Join-Path $installPath "UltraViewer_setup.exe"

    Write-Host "`n📦 Installing UltraViewer to: $installPath" -ForegroundColor Cyan
    try {
        if (-not (Test-Path $installPath)) {
            New-Item -ItemType Directory -Path $installPath -Force | Out-Null
        }
        Write-Host "⏬ Downloading UltraViewer..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $uvUrl -OutFile $exePath -UseBasicParsing -ErrorAction Stop

        Write-Host "🚀 Running UltraViewer installer silently..." -ForegroundColor Yellow
        Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT", "/NORESTART" -Wait -NoNewWindow -ErrorAction SilentlyContinue

        $installedPath = "C:\Program Files\UltraViewer\UltraViewer.exe"
        if (Test-Path $installedPath) {
            if (New-DesktopShortcut -TargetPath $installedPath -ShortcutName "UltraViewer") {
                Write-Host "✅ UltraViewer installed and shortcut created." -ForegroundColor Green
            } else {
                Write-Host "✅ UltraViewer installed. Shortcut creation failed or skipped." -ForegroundColor Yellow
            }
        } else {
            Write-Host "⚠️ UltraViewer installer ran but expected EXE not found; verify install location." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "❌ Error installing UltraViewer: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-NormalSoftware {
    Ensure-PackageManagers

    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader (Sumatra)"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="custom-anydesk"},
        @{Name="UltraViewer"; ID="custom-ultraviewer"}
    )

    Write-Host "`n📋 Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
    }

    $selection = Read-Host "Enter selection (e.g., 1,3,5). Or 'all' to install everything"
    if ($selection.Trim().ToLower() -eq 'all') {
        $selectedIndices = 1..$softwareList.Count
    } else {
        $selectedIndices = @()
        foreach ($token in ($selection -split ",")) {
            $t = $token.Trim()
            if ($t -match '^\d+$') {
                $selectedIndices += [int]$t
            }
        }
    }

    foreach ($index in $selectedIndices | Sort-Object -Unique) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            if ($app.ID -eq "custom-anydesk") {
                Install-AnyDeskDirectly
            } elseif ($app.ID -eq "custom-ultraviewer") {
                Install-UltraViewerDirectly
            } else {
                Write-Host "`n📦 Installing $($app.Name) via winget..." -ForegroundColor Gray
                try {
                    if (Get-Command winget -ErrorAction SilentlyContinue) {
                        # Use --accept flags to avoid prompts
                        Start-Process -FilePath "winget" -ArgumentList "install --id $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -ErrorAction Stop
                        Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                    } else {
                        Write-Host "⚠️ winget not available. Skipping $($app.Name)." -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "❌ Failed to install $($app.Name): $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  ⚠️ Invalid selection: $index" -ForegroundColor Red
        }
    }

    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Office installation
# -------------------------------------------------------------------------
function Install-MSOffice {
    Write-Host "`n💼 Choose MS Office installation method:" -ForegroundColor Yellow
    Write-Host "   [1] Install via Winget (simpler & faster)" -ForegroundColor White
    Write-Host "   [2] Install via Office Deployment Tool (ODT) (coming soon)" -ForegroundColor White
    Write-Host "   [0] Go Back" -ForegroundColor DarkGray
    $subChoice = Read-Host "Enter your choice [0-2]"

    switch ($subChoice) {
        '1' {
            Write-Host "`n📦 Installing Office (winget)..." -ForegroundColor Yellow
            try {
                if (Get-Command winget -ErrorAction SilentlyContinue) {
                    Start-Process -FilePath "winget" -ArgumentList "install --id Microsoft.Office.LTSC2021.ProPlus --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -ErrorAction Stop
                    Write-Host "✅ Office installed successfully (via winget)!" -ForegroundColor Green
                } else {
                    Write-Host "❌ winget not available. Cannot install Office via winget." -ForegroundColor Red
                }
            } catch {
                Write-Host "❌ Failed to install Office: $($_.Exception.Message)" -ForegroundColor Red
            }
            Read-Host "Press Enter to return to the menu..."
        }
        '2' {
            # Placeholder for ODT flow (implement extractor, fixed path handling)
            Write-Host "`n🚧 Office Deployment Tool integration is under development." -ForegroundColor Yellow
            Show-Loading -Message "Showing status"
            Read-Host "Press Enter to return to the menu..."
        }
        default {
            return
        }
    }
}

# -------------------------------------------------------------------------
# Update all software
# -------------------------------------------------------------------------
function Update-AllSoftware {
    Ensure-PackageManagers
    Write-Host "`n🔄 Checking for software updates via winget..." -ForegroundColor Yellow
    try {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -ErrorAction Stop
            Write-Host "✅ All installed software updated successfully (winget)." -ForegroundColor Green
        } else {
            Write-Host "❌ winget not available. Cannot perform updates." -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ Failed to update some software: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Activation (remote)
# -------------------------------------------------------------------------
function Invoke-Activation {
    Write-Host "`n🔑 Running System Activation Toolkit (remote)... " -ForegroundColor Yellow
    Write-Host "⚠️ This will execute a remote script. Ensure you trust the source before proceeding." -ForegroundColor Magenta
    $continue = Read-Host "Type YES to continue or anything else to cancel"
    if ($continue -ne 'YES') {
        Write-Host "Cancelled activation." -ForegroundColor Yellow
        Read-Host "Press Enter to return to the menu..."
        return
    }
    try {
        # NOTE: Remote execution is potentially dangerous. Keep it as-is per original but wrapped.
        Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
    } catch {
        Write-Host "❌ Activation script failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Advanced Toolkit (remote)
# -------------------------------------------------------------------------
function Invoke-AdvancedToolkit {
    Write-Host "`n🚀 Running Advanced Toolkit (remote script)..." -ForegroundColor Yellow
    Write-Host "⚠️ This will execute a remote script (irm https://christitus.com/win | iex)." -ForegroundColor Magenta
    $continue = Read-Host "Type YES to continue or anything else to cancel"
    if ($continue -ne 'YES') {
        Write-Host "Cancelled Advanced Toolkit." -ForegroundColor Yellow
        Read-Host "Press Enter to return to the menu..."
        return
    }
    try {
        irm https://christitus.com/win | iex
        Write-Host "✅ Advanced Toolkit execution finished." -ForegroundColor Green
    } catch {
        Write-Host "❌ Advanced Toolkit failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Main program
# -------------------------------------------------------------------------
# Ensure package managers checked at start (non-fatal)
Ensure-PackageManagers

do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-5]"
    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '5' { Invoke-AdvancedToolkit }
        '0' {
            Write-Host "`n👋 Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            break
        }
        default {
            Show-Menu -StatusMessage "⚠️ Invalid selection! Please choose between 0-5." -StatusColor "Red"
        }
    }
} while ($true)
