<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation

.NOTES
  - Work in progress.
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
    
    # Display the .NOTES section here
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "                Work in Progress Note                      " -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " - Work in progress." -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    
    if ($StatusMessage) {
        Write-Host "[STATUS] $StatusMessage" -ForegroundColor $StatusColor
        Write-Host ""
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
    
    Show-Menu -StatusMessage "✅ Selected software installed successfully!" -StatusColor "Green"
}

function Install-MSOffice {
    try {
        Write-Host "`nDownloading MS Office setup..." -ForegroundColor Yellow
        Write-Host "`nProceeding with default MS Office setup..." -ForegroundColor Yellow
        Start-Process "msiexec.exe" -ArgumentList "/i", "C:\Path\To\OfficeSetup.msi" -Wait
        Show-Menu -StatusMessage "✅ MS Office installed successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Office installation failed: $_" -StatusColor "Red"
    }
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        Show-Menu -StatusMessage "✅ Activation completed successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Activation failed: $_" -StatusColor "Red"
    }
}

function Update-AllSoftware {
    try {
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        Show-Menu -StatusMessage "✅ All software updated successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Update failed: $_" -StatusColor "Red"
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
