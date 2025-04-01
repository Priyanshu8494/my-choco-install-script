<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation

-NOTES
  - Work in progress: MS Office installation is currently not functioning. Please ensure the setup path is correct and that the installer is available at "C:\Path\To\OfficeSetup.msi" before running the script.
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
        Write-Host "[STATUS] $StatusMessage" -ForegroundColor $StatusColor
    }

    Write-Host "Main Menu Options:" -ForegroundColor Green
    Write-Host "1. Install Essential Software" -ForegroundColor White
    Write-Host "2. Install MS Office Suite" -ForegroundColor White
    Write-Host "3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "4. Update All Installed Software (Using Winget)" -ForegroundColor White
    Write-Host "0. Exit" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Install-NormalSoftware {
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="AnyDeskSoftwareGmbH.AnyDesk"},
        @{Name="UltraViewer"; ID="UltraViewer.UltraViewer"}
    )
    
    Write-Host "Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
    }
    
    $selection = Read-Host "Enter selection (e.g., 1,3,5)"
    $selectedIndices = $selection -split "," | ForEach-Object {$_ -as [int]}
    
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            Write-Host "  Installing $($app.Name)..." -ForegroundColor Gray
            winget install --id=$($app.ID) --silent --accept-source-agreements --accept-package-agreements
        } else {
            Write-Host "  Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    Write-Host "`n✅ Selected software installed successfully!" -ForegroundColor Green
    Read-Host "`nPress Enter to return to the menu..."
}

function Install-MSOffice {
    try {
        Write-Host "`nDownloading MS Office setup..." -ForegroundColor Yellow
        Write-Host "`nProceeding with default MS Office setup..." -ForegroundColor Yellow
        Start-Process "msiexec.exe" -ArgumentList "/i", "C:\Path\To\OfficeSetup.msi" -Wait
        Write-Host "`n✅ MS Office installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Office installation failed: $_" -ForegroundColor Red
    }
    Read-Host "`nPress Enter to return to the menu..."
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        Write-Host "`n✅ Activation completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Activation failed: $_" -ForegroundColor Red
    }
    Read-Host "`nPress Enter to return to the menu..."
}

function Update-AllSoftware {
    try {
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        Write-Host "`n✅ All software updated successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Update failed: $_" -ForegroundColor Red
    }
    Read-Host "`nPress Enter to return to the menu..."
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
