<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation

.NOTES
  - Work in progress.
#>

function Ensure-PackageManagers {
    # Ensure Winget is installed and available
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "❌ Winget is not installed or not working properly!" -ForegroundColor Red
        Write-Host "Please manually install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        Read-Host "Press Enter to exit..."
        exit
    }

    # Ensure Chocolatey is installed
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "Chocolatey is not installed. Installing now..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    }
}

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "          Priyanshu Suryavanshi PC Setup Toolkit             " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    param (
        [string]$StatusMessage = "", 
        [string]$StatusColor = "Yellow"
    )
    
    Show-Header

    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "                    Work in Progress                       " -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " - Work in progress." -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan

    if ($StatusMessage) {
        Write-Host "[STATUS] $StatusMessage" -ForegroundColor $StatusColor
    }

    Write-Host ""
    Write-Host " Main Menu Options: " -ForegroundColor Green
    Write-Host " ====================" -ForegroundColor Green
    Write-Host " 1. Install Essential Software" -ForegroundColor White
    Write-Host " 2. Install MS Office Suite" -ForegroundColor White
    Write-Host " 3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host " 4. Update All Installed Software (Using Winget)" -ForegroundColor White
    Write-Host " 0. Exit" -ForegroundColor Red
    Write-Host " Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
    Write-Host " ====================" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Install-AnyDeskDirectly {
    $anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\\AnyDesk"
    $exePath = "$installPath\\AnyDesk.exe"

    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan

    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath

    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "--install `\"$installPath`\" --start-with-win --silent" -Wait

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

    Write-Host "`n📦 Installing UltraViewer to: $installPath" -ForegroundColor Cyan

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

    Write-Host "Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
    }

    $selection = Read-Host "Enter selection (e.g., 1,3,5)"
    $selectedIndices = $selection -split "," | ForEach-Object {$_ -as [int]}

    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            if ($app.ID -eq "custom-anydesk") {
                Install-AnyDeskDirectly
            } elseif ($app.ID -eq "custom-ultraviewer") {
                Install-UltraViewerDirectly
            } else {
                Write-Host "  Installing $($app.Name)..." -ForegroundColor Gray
                Start-Process -FilePath "winget" -ArgumentList "install $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow

                if ($?) {
                    Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                } else {
                    Write-Host "❌ Failed to install $($app.Name)." -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  Invalid selection: $index" -ForegroundColor Red
        }
    }

    Read-Host "nPress Enter to return to the menu..."
}

function Update-AllSoftware {
    Ensure-PackageManagers

    Write-Host "🔄 Checking for software updates via Winget..." -ForegroundColor Yellow
    Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow

    if ($?) {
        Write-Host "✅ All installed software updated successfully!" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to update some software. Please check Winget logs." -ForegroundColor Red
    }

    Read-Host "nPress Enter to return to the menu..."
}

function Invoke-Activation {
    Write-Host "Running System Activation Toolkit..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
    Read-Host "nPress Enter to return to the menu..."
}

# Main program flow
Ensure-PackageManagers  # Ensure Winget & Chocolatey are installed before proceeding

do {
    Show-Menu
    $choice = Read-Host "nEnter your choice [0-4]"

    switch ($choice) {
        '1' {
            Install-NormalSoftware
        }
        '2' {
            Install-MSOffice
        }
        '3' {
            Invoke-Activation
        }
        '4' {
            Update-AllSoftware
        }
        '0' { 
            Write-Host "Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            exit 
        }
        default {
            Show-Menu -StatusMessage "Invalid selection! Please choose between 0-4" -StatusColor "Red"
        }
    }
} while ($true)
