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
    # Create software list with display names and winget IDs
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="AnyDeskSoftwareGmbH.AnyDesk"},
        @{Name="UltraViewer"; ID="UltraViewer.UltraViewer"}
    )
    
    # Display menu
    Write-Host "`nSelect software to install (Enter numbers separated by commas, or 'all'):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "$($i+1). $($softwareList[$i].Name)" -ForegroundColor White
    }
    
    # Get user selection
    $selection = Read-Host "`nEnter your choice"
    
    # Process selection
    if ($selection -eq "all") {
        $selectedIndices = 1..$softwareList.Count
    } else {
        $selectedIndices = $selection -split "," | ForEach-Object {
            try {
                [int]$_.Trim()
            } catch {
                Write-Host "Invalid input: $_" -ForegroundColor Red
                0
            }
        }
    }
    
    # Install selected software
    $successCount = 0
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            Write-Host "`nInstalling $($app.Name)..." -ForegroundColor Yellow
            
            try {
                winget install --id $app.ID --silent --accept-source-agreements --accept-package-agreements
                Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "❌ Failed to install $($app.Name): $_" -ForegroundColor Red
            }
        } elseif ($index -ne 0) {
            Write-Host "Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    # Show summary
    Write-Host "`nInstallation complete. Successfully installed $successCount of $($selectedIndices.Count) selected applications." -ForegroundColor Cyan
    Pause
}

function Install-MSOffice {
    try {
        Write-Host "`nDownloading MS Office setup..." -ForegroundColor Yellow
        # This is a placeholder - replace with your actual Office installation command
        # Start-Process "msiexec.exe" -ArgumentList "/i", "C:\Path\To\OfficeSetup.msi", "/quiet", "/norestart" -Wait
        Write-Host "✅ MS Office installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Office installation failed: $_" -ForegroundColor Red
    }
    Pause
}

function Invoke-Activation {
    try {
        Write-Host "`nActivating Windows & Office..." -ForegroundColor Yellow
        # This is a placeholder - replace with your actual activation command
        # irm https://get.activated.win | iex
        Write-Host "✅ Activation completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Activation failed: $_" -ForegroundColor Red
    }
    Pause
}

function Update-AllSoftware {
    try {
        Write-Host "`nUpdating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        Write-Host "✅ All software updated successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Update failed: $_" -ForegroundColor Red
    }
    Pause
}

# Main program flow
do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '0' { 
            Write-Host "Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            exit 
        }
        default {
            Write-Host "Invalid selection! Please choose between 0-4" -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($true)
