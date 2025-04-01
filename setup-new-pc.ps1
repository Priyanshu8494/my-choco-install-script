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

function Install-MSOffice {
    $officeVersions = @(
        @{Name="Office 2013"; Path="C:\Path\To\Office2013Setup.msi"},
        @{Name="Office 2016"; Path="C:\Path\To\Office2016Setup.msi"},
        @{Name="Office 2019"; Path="C:\Path\To\Office2019Setup.msi"},
        @{Name="Office 2021"; Path="C:\Path\To\Office2021Setup.msi"},
        @{Name="Office 365"; Path="C:\Path\To\Office365Setup.exe"}
    )
    
    do {
        Write-Host "Select MS Office version to install:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $officeVersions.Count; $i++) {
            Write-Host "  $($i + 1). $($officeVersions[$i].Name)" -ForegroundColor White
        }

        $selection = Read-Host "Enter selection (1-5)"
        
        # Ensure the input is a valid number
        if ($selection -match "^\d+$") {
            $index = [int]$selection - 1
        } else {
            $index = -1
        }

        if ($index -ge 0 -and $index -lt $officeVersions.Count) {
            $office = $officeVersions[$index]
            Write-Host "`nInstalling $($office.Name)..." -ForegroundColor Yellow
            
            if ($office.Path -match "\.msi$") {
                Start-Process "msiexec.exe" -ArgumentList "/i", "`"$($office.Path)`"", "/quiet", "/norestart" -Wait
            } else {
                Start-Process "`"$($office.Path)`"" -ArgumentList "/silent" -Wait
            }

            Write-Host "`n✅ $($office.Name) installed successfully!" -ForegroundColor Green
            break
        } else {
            Write-Host "❌ Invalid selection! Please enter a number between 1-5." -ForegroundColor Red
        }
    } while ($true)

    Read-Host "`nPress Enter to return to the menu..."
}

# Main program flow
do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '2' {
            Install-MSOffice
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
