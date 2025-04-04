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

    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host "     Priyanshu Suryavanshi PC Setup Toolkit" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================================`n" -ForegroundColor White

    Play-TypingSound
}

function Play-TypingSound {
    $soundOptions = @(
        "C:\\Windows\\Media\\Windows Notify Calendar.wav",
        "C:\\Windows\\Media\\chimes.wav",
        "C:\\Windows\\Media\\Windows Notify Email.wav",
        "C:\\Windows\\Media\\Speech On.wav",
        "C:\\Windows\\Media\\tada.wav",
        "C:\\Windows\\Media\\Windows Ding.wav"
    )
    $soundPath = Get-Random -InputObject $soundOptions
    if (Test-Path $soundPath) {
        $player = New-Object System.Media.SoundPlayer
        $player.SoundLocation = $soundPath
        $player.Play()
    }
}

function Show-Menu {
    param (
        [string]$StatusMessage = "", 
        [string]$StatusColor = "Yellow"
    )

    Show-Header

    Write-Host "📌 Status:" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "   ➔ $StatusMessage" -ForegroundColor $StatusColor
    } else {
        Write-Host "   ➔ Ready to assist with your setup!" -ForegroundColor Green
    }

    Write-Host "`n🩰 Main Menu:" -ForegroundColor Yellow
    Write-Host "   [1] 📦 Install Essential Software" -ForegroundColor White
    Write-Host "   [2] 💼 Install MS Office Suite" -ForegroundColor White
    Write-Host "   [3] 🔑 System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "   [4] 🔄 Update All Installed Software (via Winget)" -ForegroundColor White
    Write-Host "   [0] ❌ Exit" -ForegroundColor Red

    Write-Host "`n💡 Tip: Use keys to navigate (e.g., 1 for software install)"
    Write-Host "`n============================================================" -ForegroundColor DarkCyan
    Play-TypingSound
}

function Show-Loading {
    param([string]$Message = "Processing")
    $dots = ".", "..", "..."
    for ($i = 0; $i -lt 3; $i++) {
        Write-Host "`r$Message$dots[$i]" -NoNewline
        Start-Sleep -Milliseconds 300
    }
    Write-Host "`r$Message... Done!`n" -ForegroundColor Green
    Play-TypingSound
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
        Play-TypingSound
    }
}

function Install-AnyDeskDirectly {
    $anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\\AnyDesk"
    $exePath = "$installPath\\AnyDesk.exe"

    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan
    Play-TypingSound

    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath

    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "--install `$installPath` --start-with-win --silent" -Wait

    Write-Host "✅ AnyDesk installed." -ForegroundColor Green
    Play-TypingSound

    $desktop = [Environment]::GetFolderPath("Desktop")
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$desktop\\AnyDesk.lnk")
    $Shortcut.TargetPath = $exePath
    $Shortcut.IconLocation = "$exePath,0"
    $Shortcut.Save()

    Write-Host "🔗 Shortcut created on Desktop." -ForegroundColor Green
    Play-TypingSound
}

function Install-UltraViewerDirectly {
    $uvUrl = "https://ultraviewer.net/UltraViewer_setup.exe"
    $installPath = "C:\\UltraViewer"
    $exePath = "$installPath\\UltraViewer_setup.exe"

    Write-Host "`n📦 Installing UltraViewer to: $installPath" -ForegroundColor Cyan
    Play-TypingSound

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
        Play-TypingSound
    } else {
        Write-Host "❌ UltraViewer installation may have failed." -ForegroundColor Red
    }
}

Show-Menu

while ($true) {
    $choice = Read-Host "`nEnter your choice"

    switch ($choice) {
        "1" {
            Show-Menu -StatusMessage "Installing Essential Software..." -StatusColor "Cyan"
            Play-TypingSound
            # Call your essential software install function here
        }
        "2" {
            Show-Menu -StatusMessage "Installing MS Office Suite..." -StatusColor "Cyan"
            Play-TypingSound
            # Call your MS Office install function here
        }
        "3" {
            Show-Menu -StatusMessage "Launching Activation Toolkit..." -StatusColor "Cyan"
            Play-TypingSound
            # Call your activation function here
        }
        "4" {
            Show-Menu -StatusMessage "Updating all software..." -StatusColor "Cyan"
            Play-TypingSound
            winget upgrade --all
        }
        "0" {
            Write-Host "`n👋 Exiting setup. Have a great day!" -ForegroundColor Green
            Play-TypingSound
            break
        }
        default {
            Write-Host "❌ Invalid choice. Please select a valid option (0-4)." -ForegroundColor Red
            Play-TypingSound
        }
    }
}
