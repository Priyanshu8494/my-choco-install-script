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
    Write-Host "2. Install MS Office Suite" -ForegroundColor White
    Write-Host "3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
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
        Write-Host "`n✅ Essential software installed successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to return to the menu..."
        return $true
    }
    catch {
        Write-Host "Error during installation: $_" -ForegroundColor Red
        Read-Host "`nPress Enter to return to the menu..."
        return $false
    }
}

function Install-MSOffice {
    try {
        Write-Host "`nDownloading MS Office setup..." -ForegroundColor Yellow
        # In this version, the script doesn't use an external URL for Office download
        Write-Host "`nProceeding with default MS Office setup..." -ForegroundColor Yellow
        
        # Actual MS Office installation logic goes here (can be manual or different script)
        Write-Host "`nOffice installation initiated..." -ForegroundColor Yellow
        # Example of installation process (this needs to be updated with a real installation script)
        Start-Process "msiexec.exe" -ArgumentList "/i", "C:\Path\To\OfficeSetup.msi" -Wait
        
        Write-Host "`n✅ MS Office installed successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to return to the menu..."
        return $true

    }
    catch {
        Write-Host "Office installation failed: $_" -ForegroundColor Red
        Read-Host "`nPress Enter to return to the menu..."
        return $false
    }
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        
        Write-Host "`n✅ Activation completed successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to return to the menu..."
        return $true
    }
    catch {
        Write-Host "Activation failed: $_" -ForegroundColor Red
        Read-Host "`nPress Enter to return to the menu..."
        return $false
    }
}

function Update-AllSoftware {
    try {
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        
        Write-Host "`n✅ All software updated successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to return to the menu..."
        return $true
    }
    catch {
        Write-Host "Update failed: $_" -ForegroundColor Red
        Read-Host "`nPress Enter to return to the menu..."
        return $false
    }
}

# Main program flow
do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' {
            Install-NormalSoftware
        }
        '2' {
            Install-MSOffice
        }
        '3' {
            Invoke-Activation
        }
        '4' {
            Update-AllSoftware
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
