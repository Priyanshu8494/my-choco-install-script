<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
#>

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
    Write-Host "4. Update All Installed Software (Using Winget)" -ForegroundColor White
    Write-Host "0. Exit`n" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Install-NormalSoftware {
    try {
        $software = @(
            "Google.Chrome", "Mozilla.Firefox", "RARLab.WinRAR", "VideoLAN.VLC",
            "Adobe.Acrobat.Reader.64-bit", "AnyDeskSoftwareGmbH.AnyDesk", "UltraViewer.UltraViewer"
        )
        
        Write-Host "Installing software packages..." -ForegroundColor Yellow
        foreach ($app in $software) {
            Write-Host "  Installing $app..." -ForegroundColor Gray
            winget install --id=$app --silent --accept-source-agreements --accept-package-agreements
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
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        
        Write-Host "`n✅ All software updated successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to continue..."  # 👈 Prevents auto-closing after update
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
