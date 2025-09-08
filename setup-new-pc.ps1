<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
.NOTES
  - Clean UI Edition
  - Work in progress.
#>
function Show-Header {
    Clear-Host
    Write-Host "n============================================================" -ForegroundColor White
    Write-Host "     ==/          i     i          \==" -ForegroundColor DarkGray
    Write-Host "   /XX/            |\_/|            \XX\" -ForegroundColor DarkGray
    Write-Host " /XXXX\            |XXXXX|            /XXXX\" -ForegroundColor DarkGray
    Write-Host "|XXXXXX\         _XXXXXXX         /XXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXxxxxxxxXXXXXXXXXXXxxxxxxxXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host " XXXXXX/^^^^"\XXXXXXXXXXXXXXXXXXXXX/^^^^"\XXXXXX" -ForegroundColor DarkGray
    Write-Host "  |XX|       \XXX/^^\XXXXX/^^\XXX/       |XX|" -ForegroundColor DarkGray
    Write-Host "    \|       \X/    \XXX/    \X/       |/" -ForegroundColor DarkGray
    Write-Host "             !       \X/       !" -ForegroundColor DarkGray
    Write-Host "                     !" -ForegroundColor DarkGray
    Write-Host "n============================================================" -ForegroundColor White
    Write-Host "     Priyanshu Suryavanshi PC Setup Toolkit" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================================n" -ForegroundColor White
}
function Show-Menu {
    param (
        [string]$StatusMessage = "", 
        [string]$StatusColor = "Yellow"
    )
    
    Show-Header
    Write-Host "📌 Status:" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "   ➤ $StatusMessage" -ForegroundColor $StatusColor
    } else {
        Write-Host "   ➤ Ready to assist with your setup!" -ForegroundColor Green
    }
    Write-Host "n🧰 Main Menu:" -ForegroundColor Yellow
    Write-Host "   [1] 📦 Install Essential Software" -ForegroundColor White
    Write-Host "   [2] 💼 Install MS Office Suite" -ForegroundColor White
    Write-Host "   [3] 🔑 System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "   [4] 🔄 Update All Installed Software (via Winget)" -ForegroundColor White
    Write-Host "   [5] 🚀 Advance Toolkit (runs irm christitus.com/win | iex)" -ForegroundColor White
    Write-Host "   [0] ❌ Exit" -ForegroundColor Red
    Write-Host "n💡 Tip: Use numbers to navigate (e.g., 1 for software install)"
    Write-Host "n============================================================" -ForegroundColor DarkCyan
}
function Show-Loading {
    param([string]$Message = "Processing")
    $dots = ".", "..", "..."
    for ($i = 0; $i -lt 3; $i++) {
        Write-Host "r$Message$dots[$i]" -NoNewline
        Start-Sleep -Milliseconds 300
    }
    Write-Host "r$Message... Done!n" -ForegroundColor Green
}
function Ensure-PackageManagers {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "❌ Winget is not installed or not working properly!" -ForegroundColor Red
        Write-Host "➡️ Please manually install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        Read-Host "Press any key to return to the menu..."
        exit
    }
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "🍫 Chocolatey is not installed. Installing now..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    }
}
function Install-AnyDeskDirectly {
    $anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\\AnyDesk"
    $exePath = "$installPath\\AnyDesk.exe"
    Write-Host "n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }
    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath
    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "--install $installPath --start-with-win --silent" -Wait
    Write-Host "✅ AnyDesk installed." -ForegroundColor Green
    $desktop = [Environment]::GetFolderPath("Desktop")
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$desktop\\AnyDesk.lnk")
    $Shortcut.TargetPath = $exePath
    $Shortcut.IconLocation = "$exePath,0"
    $Shortcut.Save()
    Write-Host "🔗 Shortcut created on Desktop." -ForegroundColor Green
}
function Install-UltraViewerDirectly {
    $uvUrl = "https://ultraviewer.net/UltraViewer_setup.exe"
    $installPath = "C:\\UltraViewer"
    $exePath = "$installPath\\UltraViewer_setup.exe"
    Write-Host "n📦 Installing UltraViewer to: $installPath" -ForegroundColor Cyan
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }
    Write-Host "⏬ Downloading UltraViewer..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $uvUrl -OutFile $exePath
    Write-Host "🚀 Running UltraViewer installer silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT /NORESTART" -Wait
    $installedPath = "C:\\Program Files\\UltraViewer\\UltraViewer.exe"
    if (Test-Path $installedPath) {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$desktop\\UltraViewer.lnk")
        $Shortcut.TargetPath = $installedPath
        $Shortcut.IconLocation = "$installedPath,0"
        $Shortcut.Save()
        Write-Host "✅ UltraViewer installed and shortcut created." -ForegroundColor Green
    } else {
        Write-Host "❌ UltraViewer installation may have failed." -ForegroundColor Red
    }
}
function Install-NormalSoftware {
    Ensure-PackageManagers
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="custom-anydesk"},
        @{Name="UltraViewer"; ID="custom-ultraviewer"}
    )
    Write-Host "n📋 Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
    }
    $selection = Read-Host "Enter selection (e.g., 1,3,5)"
    $selectedIndices = $selection -split "," | ForEach-Object { $.Trim() -as [int] }
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            if ($app.ID -eq "custom-anydesk") {
                Install-AnyDeskDirectly
            } elseif ($app.ID -eq "custom-ultraviewer") {
                Install-UltraViewerDirectly
            } else {
                Write-Host "n📦 Installing $($app.Name)..." -ForegroundColor Gray
                Start-Process -FilePath "winget" -ArgumentList "install $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
                if ($?) {
                    Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                } else {
                    Write-Host "❌ Failed to install $($app.Name)." -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  ⚠️ Invalid selection: $index" -ForegroundColor Red
        }
    }
    Read-Host "Press any key to return to the menu..."
}
function Install-MSOffice {
    Write-Host "n💼 Choose MS Office installation method:" -ForegroundColor Yellow
    Write-Host "   [1] Install via Winget (simpler & faster)" -ForegroundColor White
    Write-Host "   [2] Install via Office Deployment Tool (ODT)" -ForegroundColor White
    Write-Host "   [0] Go Back" -ForegroundColor DarkGray
    $subChoice = Read-Host "Enter your choice [0-2]"
    switch ($subChoice) {
        '1' {
            Write-Host "n📦 Installing Office 2021 via Winget..." -ForegroundColor Yellow
            Start-Process -FilePath "winget" -ArgumentList "install Microsoft.Office LTSC2021.ProPlus --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
            if ($?) {
                Write-Host "✅ Office installed successfully!" -ForegroundColor Green
            } else {
                Write-Host "❌ Failed to install Office. Check your internet or try again." -ForegroundColor Red
            }
            Read-Host "Press any key to return to the menu..."
        }
        '2' {
            $msg = "Coming Soon"
            for ($i = 0; $i -lt 3; $i++) {
                Write-Host "r💬 $msg" -ForegroundColor Cyan -NoNewline
                Start-Sleep -Milliseconds 300
                $msg += "."
            }
            Write-Host "r🚧 Office Deployment Tool version is under development!" -ForegroundColor Yellow
            Read-Host "Press any key to return to the menu..."
        }
        default {
            return
        }
    }
}
function Update-AllSoftware {
    Ensure-PackageManagers
    Write-Host "n🔄 Checking for software updates via Winget..." -ForegroundColor Yellow
    Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
    if ($?) {
        Write-Host "✅ All installed software updated successfully!" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to update some software. Please check Winget logs." -ForegroundColor Red
    }
    Read-Host "Press any key to return to the menu..."
}
function Invoke-Activation {
    Write-Host "n🔑 Running System Activation Toolkit..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
    Read-Host "Press any key to return to the menu..."
}
function Invoke-AdvancedToolkit {
    # This is the new “Advance Toolkit” – runs the requested command
    Write-Host "n🚀 Running Advanced Toolkit (irm christitus.com/win | iex)..." -ForegroundColor Yellow
    irm christitus.com/win | iex
    Write-Host "✅ Advanced Toolkit execution finished." -ForegroundColor Green
    Read-Host "Press any key to return to the menu..."
}
# 🟢 MAIN PROGRAM START
Ensure-PackageManagers
do {
    Show-Menu
    $choice = Read-Host "nEnter your choice [0-5]"
    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '5' { Invoke-AdvancedToolkit }
        '0' {
            Write-Host "n👋 Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            break
        }
        default {
            Show-Menu -StatusMessage "⚠️ Invalid selection! Please choose between 0-5." -StatusColor "Red"
        }
    }
} while ($true)
n============================================================" -ForegroundColor White
    Write-Host "     ==/          i     i          \==" -ForegroundColor DarkGray
    Write-Host "   /XX/            |\_/|            \XX\" -ForegroundColor DarkGray
    Write-Host " /XXXX\            |XXXXX|            /XXXX\" -ForegroundColor DarkGray
    Write-Host "|XXXXXX\         _XXXXXXX         /XXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXxxxxxxxXXXXXXXXXXXxxxxxxxXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host " XXXXXX/^^^^"\XXXXXXXXXXXXXXXXXXXXX/^^^^"\XXXXXX" -ForegroundColor DarkGray
    Write-Host "  |XX|       \XXX/^^\XXXXX/^^\XXX/       |XX|" -ForegroundColor DarkGray
    Write-Host "    \|       \X/    \XXX/    \X/       |/" -ForegroundColor DarkGray
    Write-Host "             !       \X/       !" -ForegroundColor DarkGray
    Write-Host "                     !" -ForegroundColor DarkGray
    Write-Host "n" -ForegroundColor White
}
function Show-Menu {
    param (
        [string]$StatusMessage = "", 
        [string]$StatusColor = "Yellow"
    )
    
    Show-Header
    Write-Host "📌 Status:" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "   ➤ $StatusMessage" -ForegroundColor $StatusColor
    } else {
        Write-Host "   ➤ Ready to assist with your setup!" -ForegroundColor Green
    }
    Write-Host "n💡 Tip: Use numbers to navigate (e.g., 1 for software install)"
    Write-Host "r$Message$dots[$i]" -NoNewline
        Start-Sleep -Milliseconds 300
    }
    Write-Host "n" -ForegroundColor Green
}
function Ensure-PackageManagers {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "❌ Winget is not installed or not working properly!" -ForegroundColor Red
        Write-Host "➡️ Please manually install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        Read-Host "Press any key to return to the menu..."
        exit
    }
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "🍫 Chocolatey is not installed. Installing now..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    }
}
function Install-AnyDeskDirectly {
    $anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\\AnyDesk"
    $exePath = "$installPath\\AnyDesk.exe"
    Write-Host "$installPathn📦 Installing UltraViewer to: $installPath" -ForegroundColor Cyan
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }
    Write-Host "⏬ Downloading UltraViewer..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $uvUrl -OutFile $exePath
    Write-Host "🚀 Running UltraViewer installer silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT /NORESTART" -Wait
    $installedPath = "C:\\Program Files\\UltraViewer\\UltraViewer.exe"
    if (Test-Path $installedPath) {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$desktop\\UltraViewer.lnk")
        $Shortcut.TargetPath = $installedPath
        $Shortcut.IconLocation = "$installedPath,0"
        $Shortcut.Save()
        Write-Host "✅ UltraViewer installed and shortcut created." -ForegroundColor Green
    } else {
        Write-Host "❌ UltraViewer installation may have failed." -ForegroundColor Red
    }
}
function Install-NormalSoftware {
    Ensure-PackageManagers
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="custom-anydesk"},
        @{Name="UltraViewer"; ID="custom-ultraviewer"}
    )
    Write-Host "n📦 Installing $($app.Name)..." -ForegroundColor Gray
                Start-Process -FilePath "winget" -ArgumentList "install $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
                if ($?) {
                    Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                } else {
                    Write-Host "❌ Failed to install $($app.Name)." -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  ⚠️ Invalid selection: $index" -ForegroundColor Red
        }
    }
    Read-Host "Press any key to return to the menu..."
}
function Install-MSOffice {
    Write-Host "n📦 Installing Office 2021 via Winget..." -ForegroundColor Yellow
            Start-Process -FilePath "winget" -ArgumentList "install Microsoft.Office LTSC2021.ProPlus --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
            if ($?) {
                Write-Host "✅ Office installed successfully!" -ForegroundColor Green
            } else {
                Write-Host "❌ Failed to install Office. Check your internet or try again." -ForegroundColor Red
            }
            Read-Host "Press any key to return to the menu..."
        }
        '2' {
            $msg = "Coming Soon"
            for ($i = 0; $i -lt 3; $i++) {
                Write-Host "r🚧 Office Deployment Tool version is under development!" -ForegroundColor Yellow
            Read-Host "Press any key to return to the menu..."
        }
        default {
            return
        }
    }
}
function Update-AllSoftware {
    Ensure-PackageManagers
    Write-Host "n🔑 Running System Activation Toolkit..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
    Read-Host "Press any key to return to the menu..."
}
function Invoke-AdvancedToolkit {
    # This is the new “Advance Toolkit” – runs the requested command
    Write-Host "nEnter your choice [0-5]"
    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '5' { Invoke-AdvancedToolkit }
        '0' {
            Write-Host "
