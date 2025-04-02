# Install software using Chocolatey

# Check if Chocolatey is installed
if (-not (Test-Path "C:\ProgramData\Chocolatey")) {
    Write-Host "Chocolatey is not installed. Installing now..."
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

# Install WinRAR
Write-Host "Installing WinRAR..."
choco install winrar -y

# Install Google Chrome
Write-Host "Installing Google Chrome..."
choco install googlechrome -y

# Install Firefox
Write-Host "Installing Firefox..."
choco install firefox -y

# Install VLC
Write-Host "Installing VLC..."
choco install vlc -y

Write-Host "Software installation completed."
