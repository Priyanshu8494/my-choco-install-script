# Must be run as Administrator

# ----------------- CONFIG -----------------
$installPath = "C:\UltraViewer"
$ultraViewerUrl = "https://ultraviewer.net/en/download.html"
$installerPath = "$env:TEMP\UltraViewer_Setup.exe"
$shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "UltraViewer.lnk")
# -----------------------------------------

function Ensure-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "❌ Please run this script as an Administrator." -ForegroundColor Red
        exit
    }
}

function Download-UltraViewer {
    Write-Host "`n⏬ Downloading UltraViewer..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $ultraViewerUrl -OutFile $installerPath
        Write-Host "✅ Download completed." -ForegroundColor Green
    } catch {
        Write-Host "❌ Download failed. Please check the URL or your internet connection." -ForegroundColor Red
        exit
    }
}

function Install-UltraViewer {
    Write-Host "`n📂 Installing UltraViewer to: $installPath" -ForegroundColor Cyan
    $installArgs = "/VERYSILENT /DIR=`"$installPath`""
    try {
        Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait
        Write-Host "✅ UltraViewer installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "❌ Installation failed." -ForegroundColor Red
        exit
    }
}

function Create-Shortcut {
    Write-Host "`n🔗 Creating shortcut on Desktop..." -ForegroundColor Cyan
    $targetExe = [System.IO.Path]::Combine($installPath, "UltraViewer.exe")
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetExe
        $shortcut.WorkingDirectory = $installPath
        $shortcut.IconLocation = $targetExe
        $shortcut.Save()
        Write-Host "✅ Shortcut created: $shortcutPath" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to create shortcut." -ForegroundColor Red
        exit
    }
}

# ------------------ EXECUTION ------------------
Ensure-Admin
Download-UltraViewer
Install-UltraViewer
Create-Shortcut

Write-Host "`n🎉 All tasks completed successfully! UltraViewer is ready to use." -ForegroundColor Green
