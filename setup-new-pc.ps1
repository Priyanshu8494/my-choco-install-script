# Install Chocolatey if not installed
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-Output "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    iwr https://chocolatey.org/install.ps1 -UseBasicParsing | iex
}

# Install selected apps
Write-Output "Installing Google Chrome, Firefox, WinRAR, and VLC Player..."
choco install googlechrome firefox winrar vlc -y

Write-Output "Installation complete!"
