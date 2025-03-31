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
        do {
            Write-Host "`nChoose MS Office Edition to Install:" -ForegroundColor Yellow
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
                    Read-Host "`nPress Enter to return to the menu..."
                    return $false
                }
                default {
                    Write-Host "Invalid selection! Please choose a valid option." -ForegroundColor Red
                    continue
                }
            }

            Write-Host "`nDownloading MS Office $edition installation script..." -ForegroundColor Yellow
            $officeScript = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/Priyanshu8494/ms-office-install-script/main/setup-$edition.ps1"
            
            Write-Host "Executing MS Office installation for $edition..." -ForegroundColor Yellow
            $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
            $officeScript | Out-File $tempFile
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempFile`"" -Wait
            Remove-Item $tempFile -Force
            
            Write-Host
