<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
#>

function Initialize-Chocolatey {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        try {
            Write-Host "Chocolatey not found. Installing Chocolatey package manager..." -ForegroundColor Magenta
            Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop | Out-Null
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1')) | Out-Null
            # Update PATH environment variable
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
            return $true
        }
        catch {
            Write-Host "Failed to install Chocolatey: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "            Priyanshu Suryavanshi PC Setup Toolkit          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    param ([string]$StatusMessage = "", [string]$StatusColor = "Yellow")
    
    Show-Header
    
    if ($StatusMessage) {
        Write-Host "[STATUS] $StatusMessage`n" -ForegroundColor $StatusColor
    }

    Write-Host "Main Menu Options:`n" -ForegroundColor Green
    Write-Host "1. Install Essential Software" -ForegroundColor White
    Write-Host "   (Chrome, Firefox, WinRAR, VLC, Adobe Reader, AnyDesk, UltraViewer)"
    Write-Host "2. Install MS Office Suite" -ForegroundColor White
    Write-Host "3. System Activation Toolkit" -ForegroundColor White
    Write-Host "4. Update All Installed Software" -ForegroundColor White
    Write-Host "0. Exit`n" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Install-NormalSoftware {
    try {
        if (-not (Initialize-Chocolatey)) {
            return $false
        }
        
        $software = @('googlechrome', 'firefox', 'winrar', 'vlc', 'adobereader', 'anydesk', 'ultraviewer')
        Write-Host "Installing software packages..." -ForegroundColor Yellow
        foreach ($app in $software) {
            Write-Host "  Installing $app..." -ForegroundColor Gray
            choco install $app -y --force | Out-Null
        }
        return $true
    }
    catch {
        Write-Host "Error during installation: $_" -ForegroundColor Red
        return $false
    }
}

function Install-MSOffice {
    try {
        Write-Host "Downloading Office installation script..." -ForegroundColor Yellow
        $officeScript = Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/Priyanshu8494/ms-office-install-script/main/setup-office.ps1'
        
        Write-Host "Executing installation..." -ForegroundColor Yellow
        $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
        $officeScript | Out-File $tempFile
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempFile`"" -Wait
        Remove-Item $tempFile -Force
        
        return $true
    }
    catch {
        Write-Host "Office installation failed: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-Activation {
    try {
        Write-Host "Starting activation process..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        return $true
    }
    catch {
        Write-Host "Activation failed: $_" -ForegroundColor Red
        return $false
    }
}

function Update-AllSoftware {
    try {
        if (-not (Initialize-Chocolatey)) {
            return $false
        }
        
        Write-Host "Updating all installed software..." -ForegroundColor Yellow
        choco upgrade all -y --ignore-checksums | Out-Null
        Write-Host "✅ All software updated successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Update failed: $_" -ForegroundColor Red
        return $false
    }
}

# Main program flow
do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' {
            if (Install-NormalSoftware) {
                Show-Menu -StatusMessage "Essential software installed successfully!" -StatusColor "Green"
            } else {
                Show-Menu -StatusMessage "Software installation failed. Check permissions and internet connection." -StatusColor "Red"
            }
        }
        '2' {
            if (Install-MSOffice) {
                Show-Menu -StatusMessage "Office suite installed successfully!" -StatusColor "Green"
            } else {
                Show-Menu -StatusMessage "Office installation failed. Verify repository access." -StatusColor "Red"
            }
        }
        '3' {
            if (Invoke-Activation) {
                Show-Menu -StatusMessage "System activation completed successfully!" -StatusColor "Green"
            } else {
                Show-Menu -StatusMessage "Activation process failed. Try running as administrator." -StatusColor "Red"
            }
        }
        '4' {
            if (Update-AllSoftware) {
                Show-Menu -StatusMessage "All software updated successfully!" -StatusColor "Green"
            } else {
                Show-Menu -StatusMessage "Update failed. Please check Chocolatey installation." -StatusColor "Red"
            }
        }
        '0' { 
            Write-Host "Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            exit 
        }
        default {
            Show-Menu -StatusMessage "Invalid selection! Please choose between 0-4" -StatusColor "Red"
        }
    }
} while ($true)
