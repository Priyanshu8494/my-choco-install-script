Clear-Host

# Center the PowerShell window
$hwnd = (Get-Process -Id $pid).MainWindowHandle
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Window {
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hwnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hwnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
}
"@
[Window]::SetWindowPos($hwnd, [IntPtr]::Zero, [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Width / 4, [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Height / 4, 0, 0, 0)

function Show-Menu {
    param (
        [string]$Message = ""
    )

    if ($Message -ne "") {
        Write-Host "`n$Message`n" -ForegroundColor Yellow
    }

    Write-Host "      _==/--" -ForegroundColor Yellow
    Write-Host "     /##/     |\_/|     \##\" -ForegroundColor Yellow
    Write-Host "    |####|    (o.o)    |####|" -ForegroundColor Yellow
    Write-Host "     \##\     > ^ <     /##/" -ForegroundColor Yellow
    Write-Host "      '--'               '--'" -ForegroundColor Yellow

    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "          PC SETUP [ PRIYANSHU SURYAVANSHI ]           " -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan

    Write-Host "[1] Install Normal Software (Chrome, WinRAR, VLC, Firefox, Adobe Reader, AnyDesk, UltraViewer)" -ForegroundColor White
    Write-Host "[2] Install MS Office (via GitHub)" -ForegroundColor White
    Write-Host "[3] Activation Tool (Using Custom URL)" -ForegroundColor White
    Write-Host "[0] Exit" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Cyan

    $Key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $choice = $Key.Character

    switch ($choice) {
        '1' {
            Write-Host "Installing Normal Software..." -ForegroundColor Yellow
            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                Write-Host "Chocolatey not found. Installing Chocolatey..." -ForegroundColor Magenta
                Set-ExecutionPolicy Bypass -Scope Process -Force
                iwr https://chocolatey.org/install.ps1 -UseBasicParsing | iex
            }
            choco install googlechrome firefox winrar vlc adobereader anydesk ultraviewer -y
            Show-Menu "✅ Installation complete!"
        }
        '2' {
            Write-Host "Installing MS Office..." -ForegroundColor Yellow
            powershell -ep Bypass -c "iwr 'https://raw.githubusercontent.com/Priyanshu8494/ms-office-install-script/main/setup-office.ps1' | iex"
            Show-Menu "✅ MS Office Installation complete!"
        }
        '3' {
            Write-Host "Running Activation Tool..." -ForegroundColor Yellow
            irm https://get.activated.win | iex
            Show-Menu "✅ Activation process completed!"
        }
        '0' {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Red; exit
        }
        default {
            Show-Menu "❌ Invalid selection. Please try again."
        }
    }
}

Show-Menu
