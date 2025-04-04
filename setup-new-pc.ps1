# Must be run as Administrator

# ------------- CONFIG -------------
$installPath = "C:\UltraViewer"
$ultraViewerUrl = "https://www.ultraviewer.net/UltraViewer_setup_6.6_en.exe"
$ultraViewerInstaller = "$env:TEMP\UltraViewerInstaller.exe"
$shortcutPath = "$env:PUBLIC\Desktop\UltraViewer.lnk"
# ----------------------------------

function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Please run this script as Administrator." -ForegroundColor Red
        exit
    }
}

function Install-UltraViewer {
    Write-Host "`n📦 Downloading UltraViewer..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $ultraViewerUrl -OutFile $ultraViewerInstaller

    Write-Host "📂 Installing to: $installPath" -ForegroundColor Cyan
    Start-Process -FilePath $ultraViewerInstaller -ArgumentList "/VERYSILENT /DIR=`"$installPath`"" -Wait

    Write-Host "✅ UltraViewer installed successfully!" -ForegroundColor Green
}

function Create-Shortcut {
    Write-Host "🔗 Creating shortcut on Desktop..." -ForegroundColor Cyan

    $targetExe = Join-Path $installPath "UltraViewer.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetExe
    $shortcut.WorkingDirectory = $installPath
    $shortcut.IconLocation = $targetExe
    $shortcut.Save()

    Write-Host "✅ Shortcut created: $shortcutPath" -ForegroundColor Green
}

# ------------ EXECUTE ------------
Ensure-Admin
Install-UltraViewer
Create-Shortcut

Write-Host "`n🎉 All done! UltraViewer is ready to use." -ForegroundColor Green
