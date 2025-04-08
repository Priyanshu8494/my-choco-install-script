<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
.NOTES
  - Clean UI Edition
  - Fully Functional
#>

function Show-Header {
    Clear-Host
    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host "     _==/          i     i          \==_" -ForegroundColor DarkGray
    Write-Host "   /XX/            |\___/|            \XX\" -ForegroundColor DarkGray
    Write-Host " /XXXX\            |XXXXX|            /XXXX\" -ForegroundColor DarkGray
    Write-Host "|XXXXXX\_         _XXXXXXX_         _/XXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXxxxxxxxXXXXXXXXXXXxxxxxxxXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" -ForegroundColor DarkGray
    Write-Host "|XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX|" -ForegroundColor DarkGray
    Write-Host " XXXXXX/^^^^\XXXXXXXXXXXXXXXXXXXXX/^^^^\XXXXXX" -ForegroundColor DarkGray
    Write-Host "  |XX|       \XXX/^^\XXXXX/^^\XXX/       |XX|" -ForegroundColor DarkGray
    Write-Host "    \|       \X/    \XXX/    \X/       |/" -ForegroundColor DarkGray
    Write-Host "             !       \X/       !" -ForegroundColor DarkGray
    Write-Host "                     !" -ForegroundColor DarkGray
    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host "     Priyanshu Suryavanshi PC Setup Toolkit" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================================`n" -ForegroundColor White
}

function Ensure-PackageManagers {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "⚠️ Winget not found. Please install App Installer from Microsoft Store." -ForegroundColor Red
        Pause
        Exit
    }
}

function Install-AnyDeskDirectly {
    $url = "https://download.anydesk.com/AnyDesk.exe"
    $installer = "$env:TEMP\AnyDesk.exe"
    Invoke-WebRequest -Uri $url -OutFile $installer
    Start-Process -FilePath $installer -ArgumentList "/silent" -Wait
    Remove-Item $installer -Force
    Write-Host "✅ AnyDesk installed successfully." -ForegroundColor Green
}

function Install-UltraViewerDirectly {
    $url = "https://www.ultraviewer.net/en/UltraViewer_setup.exe"
    $installer = "$env:TEMP\UltraViewer_setup.exe"
    Invoke-WebRequest -Uri $url -OutFile $installer
    Start-Process -FilePath $installer -ArgumentList "/VERYSILENT" -Wait
    Remove-Item $installer -Force
    Write-Host "✅ UltraViewer installed successfully." -ForegroundColor Green
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
        @{Name="UltraViewer"; ID="custom-ultraviewer"},
        @{Name="🔙 Back to Main Menu"; ID="back"}
    )

    do {
        Write-Host "`n📋 Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
        for ($i = 0; $i -lt $softwareList.Count; $i++) {
            Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
        }

        $selection = Read-Host "Enter selection (e.g., 1,3,5)"
        if ($selection -eq "8") { return }

        $selectedIndices = $selection -split "," | ForEach-Object { $_.Trim() -as [int] }

        foreach ($index in $selectedIndices) {
            if ($index -ge 1 -and $index -le $softwareList.Count) {
                $app = $softwareList[$index - 1]
                if ($app.ID -eq "custom-anydesk") {
                    Install-AnyDeskDirectly
                } elseif ($app.ID -eq "custom-ultraviewer") {
                    Install-UltraViewerDirectly
                } elseif ($app.ID -eq "back") {
                    return
                } else {
                    Write-Host "`n📦 Installing $($app.Name)..." -ForegroundColor Gray
                    Start-Process -FilePath "winget" -ArgumentList "install $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
                    if ($?) {
                        Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                    } else {
                        Write-Host "❌ Failed to install $($app.Name)." -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "⚠️ Invalid selection: $index" -ForegroundColor Red
            }
        }
    } while ($true)
}

function Install-MSOffice {
    Clear-Host
    Show-Header
    Write-Host "`n💼 Office Installation Options:" -ForegroundColor Yellow
    Write-Host "  [1] Install via Winget (Microsoft 365 Apps)" -ForegroundColor White
    Write-Host "  [2] Coming Soon: Office Deployment Tool (ODT)" -ForegroundColor DarkGray
    Write-Host "  [0] Back to Main Menu" -ForegroundColor Red
    $choice = Read-Host "Select option"

    switch ($choice) {
        "1" {
            Write-Host "`n📦 Installing Microsoft 365 Apps..." -ForegroundColor Gray
            Start-Process -FilePath "winget" -ArgumentList "install Microsoft.Office --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
            if ($?) {
                Write-Host "✅ Microsoft Office installed successfully!" -ForegroundColor Green
            } else {
                Write-Host "❌ Installation failed." -ForegroundColor Red
            }
            Pause
        }
        "2" {
            Write-Host "`n🛠 ODT support is under construction..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }
        "0" { return }
        default {
            Write-Host "⚠️ Invalid input." -ForegroundColor Red
        }
    }
}

function Activate-System {
    Write-Host "`n🔐 Starting Activation Toolkit..." -ForegroundColor Yellow
    Write-Host "➡️  This feature is under construction." -ForegroundColor DarkGray
    Start-Sleep -Seconds 2
}

function Update-AllSoftware {
    Ensure-PackageManagers
    Write-Host "`n🔄 Checking for updates..." -ForegroundColor Gray
    Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
    if ($?) {
        Write-Host "✅ All software updated successfully!" -ForegroundColor Green
    } else {
        Write-Host "❌ Update process failed." -ForegroundColor Red
    }
    Pause
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

    Write-Host "`n🧰 Main Menu:" -ForegroundColor Yellow
    Write-Host "   [1] 📦 Install Essential Software" -ForegroundColor White
    Write-Host "   [2] 💼 Install MS Office Suite" -ForegroundColor White
    Write-Host "   [3] 🔑 System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "   [4] 🔄 Update All Installed Software (via Winget)" -ForegroundColor White
    Write-Host "   [0] ❌ Exit" -ForegroundColor Red
    Write-Host "`n💡 Tip: Use numbers to navigate (e.g., 1 for software install)"
    Write-Host "`n============================================================" -ForegroundColor DarkCyan
}

# === MAIN LOOP ===
do {
    Show-Menu
    $input = Read-Host "Enter your choice"
    switch ($input) {
        "1" { Install-NormalSoftware }
        "2" { Install-MSOffice }
        "3" { Activate-System }
        "4" { Update-AllSoftware }
        "0" { break }
        default { Write-Host "⚠️ Invalid choice. Please try again." -ForegroundColor Red }
    }
} while ($true)

Write-Host "`n👋 Thank you for using Priyanshu's Setup Toolkit. Goodbye!" -ForegroundColor Green
