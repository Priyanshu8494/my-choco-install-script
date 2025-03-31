Clear-Host

# Center the PowerShell window
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
$hwnd = (Get-Process -Id $pid).MainWindowHandle
[Window]::SetWindowPos($hwnd, [IntPtr]::Zero, [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Width / 4, [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Height / 4, 0, 0, 0)

function Show-Menu {
    param (
        [string]$Message = ""
    )

    # Suppress unwanted output and keep the window clean
    $null = $Message

    # ASCII Art (On top of the name Priyanshu Suryavanshi)
    Write-Host "        _==/          i     i          /==_ " -ForegroundColor Yellow
    Write-Host "       /##/           |\_/|           \##\" -ForegroundColor Yellow
    Write-Host "      |####|          (o.o)          |####|" -ForegroundColor Yellow
    Write-Host "       \##\           > ^ <           /##/" -ForegroundColor Yellow
    Write-Host "        '--'                         '--'" -ForegroundColor Yellow

    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "          PC SETUP [ PRIYANSHU SURYAVANSHI ]        " -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "`n"  # spacing between header and menu

    # Menu with improved colors and structure
    Write-Host "  [1] Install Normal Software: Chrome, WinRAR, VLC, Firefox, Adobe Reader, AnyDesk, UltraViewer" -ForegroundColor White
    Write-Host "  [2] Install MS Office (via GitHub)" -ForegroundColor White
    Write-Host "  [3] Run Activation Tool (Using Custom URL)" -ForegroundColor White
    Write-Host "  [0] Exit" -ForegroundColor Red
    Write-Host "`n=============================================" -ForegroundColor Cyan

    # Read the key press without requiring Enter
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").Character

    # Switch based on the key press
    switch ($key) {
        '1' {
            Write-Host "`nInstalling Normal Software..." -ForegroundColor Yellow
            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                Write-Host "Chocolatey not found. Installing Chocolatey..." -ForegroundColor Magenta
                Set-ExecutionPolicy Bypass -Scope Process -Force
                iwr https://chocolatey.org/install.ps1 -UseBasicParsing | iex | Out-Null
            }
            choco install googlechrome firefox winrar vlc adobereader anydesk ultraviewer -y | Out-Null
            Show-Menu "✅ Installation complete!"
        }
        '2' {
            Write-Host "`nInstalling MS Office..." -ForegroundColor Yellow
            powershell -
