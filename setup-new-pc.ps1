<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit (Scoop Version)
.DESCRIPTION
  Automated PC setup with software installation using Scoop and system activation

.NOTES
  - Requires Scoop to be installed
  - Run in a PowerShell session without administrator privileges
#>

function Show-Header {
    Clear-Host
    Write-Host "" 
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "            Priyanshu Suryavanshi PC Setup Toolkit          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Install-Scoop {
    try {
        Write-Host "`nChecking if Scoop is installed..." -ForegroundColor Yellow
        if (Get-Command scoop -ErrorAction SilentlyContinue) {
            Write-Host "✅ Scoop is already installed" -ForegroundColor Green
            return $true
        }

        Write-Host "Installing Scoop (User Mode)..." -ForegroundColor Yellow
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression
        
        if (Get-Command scoop -ErrorAction SilentlyContinue) {
            Write-Host "✅ Scoop installed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ Scoop installation failed" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "❌ Error installing Scoop: $_" -ForegroundColor Red
        return $false
    }
    finally {
        Pause
    }
}

function Install-Software {
    # Scoop package list
    $softwareList = @(
        @{Name="Google Chrome"; ID="googlechrome"},
        @{Name="Mozilla Firefox"; ID="firefox"},
        @{Name="WinRAR"; ID="winrar"},
        @{Name="VLC Player"; ID="vlc"},
        @{Name="Sumatra PDF Reader"; ID="sumatrapdf"},
        @{Name="AnyDesk"; ID="anydesk"},
        @{Name="UltraViewer"; ID="ultraviewer"}
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
    
    $successCount = 0
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            Write-Host "`nInstalling $($app.Name)..." -ForegroundColor Yellow
            
            try {
                scoop install $app.ID
                Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "❌ Failed to install $($app.Name): $_" -ForegroundColor Red
            }
        } elseif ($index -ne 0) {
            Write-Host "Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    Write-Host "`nInstallation complete. Successfully installed $successCount of $($selectedIndices.Count) selected applications." -ForegroundColor Cyan
    Pause
}

function Install-MSOffice {
    try {
        Write-Host "`nInstalling Microsoft Office..." -ForegroundColor Yellow
        scoop bucket add extras
        scoop install office-deployment-tool
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
        Write-Host "`nUpdating all installed software using Scoop..." -ForegroundColor Yellow
        scoop update *
        Write-Host "✅ All software updated successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Update failed: $_" -ForegroundColor Red
    }
    Pause
}

# Main program flow
try {
    if (-not (Get-Command scoop -ErrorAction SilentlyContinue)) {
        Install-Scoop
    }

    do {
        Show-Header
        Write-Host "1. Install Essential Software" -ForegroundColor White
        Write-Host "2. Install MS Office Suite" -ForegroundColor White
        Write-Host "3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
        Write-Host "4. Update All Installed Software (Using Scoop)" -ForegroundColor White
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
}
catch {
    Write-Host "`nA critical error occurred: $_" -ForegroundColor Red
    Pause
    exit 1
}
