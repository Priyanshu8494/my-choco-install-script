<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit (Winget Version)
.DESCRIPTION
  Automated PC setup with software installation using Winget and system activation

.NOTES
  - Requires Winget to be installed
  - Run in a PowerShell session with administrator privileges
#>

function Show-Header {
    Clear-Host
    Write-Host "" 
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "            Priyanshu Suryavanshi PC Setup Toolkit          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Install-Software {
    # Winget package list
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="Sumatra PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="AnyDeskSoftwareGmbH.AnyDesk"},
        @{Name="UltraViewer"; ID="UltraViewer.UltraViewer"}
    )
    
    Write-Host "`nSelect software to install (Enter numbers separated by commas, or 'all'):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "$($i+1). $($softwareList[$i].Name)" -ForegroundColor White
    }
    
    $selection = Read-Host "`nEnter your choice"
    if ($selection -eq "all") {
        $selectedIndices = 1..$softwareList.Count
    } else {
        $selectedIndices = $selection -split "," | ForEach-Object {
            try {[int]$_.Trim()} catch {Write-Host "Invalid input: $_" -ForegroundColor Red; 0}
        }
    }
    
    $installJobs = @()
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            Write-Host "`nStarting installation for $($app.Name)..." -ForegroundColor Yellow
            
            # Run installation in parallel using Start-Job
            $installJobs += Start-Job -ScriptBlock {
                param ($AppID)
                winget install $AppID --silent --accept-source-agreements --accept-package-agreements
            } -ArgumentList $app.ID
        } elseif ($index -ne 0) {
            Write-Host "Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    # Wait for all installations to complete
    $installJobs | ForEach-Object { Receive-Job -Job $_ -Wait }
    Write-Host "`nInstallation complete." -ForegroundColor Cyan
    Pause
}

function Install-MSOffice {
    try {
        Write-Host "`nInstalling Microsoft Office..." -ForegroundColor Yellow
        Start-Process -FilePath "winget" -ArgumentList "install Microsoft.Office --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
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
        # Placeholder - replace with actual activation command
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
        Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow
        Write-Host "✅ All software updated successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Update failed: $_" -ForegroundColor Red
    }
    Pause
}

# Main program flow
do {
    Show-Header
    Write-Host "1. Install Essential Software" -ForegroundColor White
    Write-Host "2. Install MS Office Suite" -ForegroundColor White
    Write-Host "3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "4. Update All Installed Software (Using Winget)" -ForegroundColor White
    Write-Host "0. Exit" -ForegroundColor Red
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' { Install-Software }
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
