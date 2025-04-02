@echo off
echo Installing software using Chocolatey...

:: Check if Chocolatey is installed
if not exist "C:\ProgramData\Chocolatey" (
    echo Chocolatey is not installed. Installing now...
    @"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;C:\ProgramData\Chocolatey\bin"
)

:: Install WinRAR
choco install winrar -y

:: Install Google Chrome
choco install googlechrome -y

:: Install Firefox
choco install firefox -y

:: Install VLC
choco install vlc -y

echo Software installation completed.
pause
