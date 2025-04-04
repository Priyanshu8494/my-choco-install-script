# Must be run as Administrator

# ----------------- CONFIG -----------------
$anydeskUrl = "http://download.anydesk.com/AnyDesk.exe"
$installPath = "C:\AnyDesk"
$exePath = "$installPath\AnyDesk.exe"
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
    Write-Host "`n📦 Installing AnyDesk to: $installPath" -ForegroundColor Cyan

    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath | Out-Null
    }

    Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath

    Write-Host "🚀 Installing AnyDesk silently..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath -ArgumentList "--install `"$installPath`" --start-with-win --silent" -Wait

    Write-Host "✅ AnyDesk installed." -ForegroundColor Green

    Create-Shortcut -exePath $exePath -shortcutPath "$(Get-DesktopPath)\AnyDesk.lnk"
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

# ------------------ EXECUTION ------------------
Ensure-Admin
Install-AnyDesk

Write-Host "`n🎉 All tasks completed successfully!" -ForegroundColor Green
