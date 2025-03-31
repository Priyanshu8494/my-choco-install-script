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
    Write-Host "2. Install MS Office Suite (Choose Edition)" -ForegroundColor White
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
        return $true
    }
    catch {
        Write-Host "Error during installation: $_" -ForegroundColor Red
        return $false
    }
}

function Install-MSOffice {
    try {
        Write-Host "Choose MS Office Edition to Install:" -ForegroundColor Yellow
        Write-Host "1. Microsoft Office 2013" -ForegroundColor White
        Write-Host "2. Microsoft Office 2019" -ForegroundColor White
        Write-Host "3. Microsoft Office 2021" -ForegroundColor White
        Write-Host "4. Microsoft Office 2024" -ForegroundColor White
        Write-Host "5. Microsoft 365 (Subscription-based)" -ForegroundColor White
        Write-Host "0. Cancel Installation" -ForegroundColor Red

        $officeChoice = Read-Host "`nEnter your choice [0-5]"

        switch ($officeChoice) {
            '1' { $edition = "Office2013" }
            '2' { $edition = "Office2019" }
            '3' { $edition = "Office2021" }
            '4' { $edition = "Office2024" }
            '5' { $edition = "Microsoft365" }
            '0' { 
                Write-Host "Office installation cancelled." -ForegroundColor Cyan
                return $false
            }
            default {
                Write-Host "Invalid selection! Cancelling installation..." -ForegroundColor Red
                return $false
            }
        }

        Write-Host "`nDownloading MS Office $edition installation script..." -ForegroundColor Yellow
        $officeScript = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/Priyanshu8494/ms-office-install-script/main/setup-$edition.ps1"
        
        Write-Host "Executing MS Office installation for $edition..." -ForegroundColor Yellow
        $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
        $officeScript | Out-File $tempFile
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempFile`"" -Wait
        Remove-Item $tempFile -Force
        
        Write-Host "`n✅ MS Office $edition installed successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to continue..."
        return $true
    }
    catch {
        Write-Host "Office installation failed: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        
        Write-Host "`n✅ Activation completed successfully!" -ForegroundColor Green
        Read-Host "`nPress Enter to continue..."
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
        Read-Host "`nPress Enter to continue..."
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
                Show-Menu -StatusMessage "MS Office installed successfully!" -StatusColor "Green"
            } else {
                Show-Menu -StatusMessage "MS Office installation failed. Verify repository access." -StatusColor "Red"
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
                Show-Menu -StatusMessage "Update failed. Ensure Winget is installed and running." -StatusColor "Red"
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
