<#
.SYNOPSIS
Automated PC setup with software installation and Office download.
.DESCRIPTION
Script to automate PC setup by installing software like Google Chrome, Firefox, WinRAR, VLC Player, and also download Office versions (2019, 2021, 2024, and 365).
#>

function Show-OfficeVersionMenu {
    Write-Host "Select Office version to download:" -ForegroundColor Green
    Write-Host "1. Office 2019" -ForegroundColor White
    Write-Host "2. Office 2021" -ForegroundColor White
    Write-Host "3. Office 2024" -ForegroundColor White
    Write-Host "4. Office 365" -ForegroundColor White
    Write-Host "0. Back to Menu" -ForegroundColor Red
    $versionChoice = Read-Host "Enter your choice [0-4]"

    switch ($versionChoice) {
        '1' { return "ProPlus2019Retail" }
        '2' { return "ProPlus2021Volume" }
        '3' { return "ProPlus2024Volume" }
        '4' { return "O365ProPlusRetail" }
        '0' { return $null }
        default { Write-Host "Invalid selection, please choose between 0-4." -ForegroundColor Red; return Show-OfficeVersionMenu }
    }
}

function Show-ChannelMenu {
    Write-Host "Select Office channel:" -ForegroundColor Green
    Write-Host "1. Current" -ForegroundColor White
    Write-Host "2. Perpetual VL 2021" -ForegroundColor White
    Write-Host "3. Perpetual VL 2024" -ForegroundColor White
    Write-Host "4. Semi-Annual" -ForegroundColor White
    Write-Host "0. Back to Menu" -ForegroundColor Red
    $channelChoice = Read-Host "Enter your choice [0-4]"

    switch ($channelChoice) {
        '1' { return "Current" }
        '2' { return "PerpetualVL2021" }
        '3' { return "PerpetualVL2024" }
        '4' { return "SemiAnnual" }
        '0' { return $null }
        default { Write-Host "Invalid selection, please choose between 0-4." -ForegroundColor Red; return Show-ChannelMenu }
    }
}

function Show-ComponentsMenu {
    Write-Host "Select Office components to download (separate by commas):" -ForegroundColor Green
    Write-Host "1. Access" -ForegroundColor White
    Write-Host "2. OneDrive" -ForegroundColor White
    Write-Host "3. Outlook" -ForegroundColor White
    Write-Host "4. Word" -ForegroundColor White
    Write-Host "5. Excel" -ForegroundColor White
    Write-Host "6. OneNote" -ForegroundColor White
    Write-Host "7. Publisher" -ForegroundColor White
    Write-Host "8. PowerPoint" -ForegroundColor White
    Write-Host "9. Teams" -ForegroundColor White
    Write-Host "10. Project 2019" -ForegroundColor White
    Write-Host "11. Project 2021" -ForegroundColor White
    Write-Host "12. Project 2024" -ForegroundColor White
    Write-Host "0. Back to Menu" -ForegroundColor Red
    $componentsChoice = Read-Host "Enter your choice [0-12]"

    switch ($componentsChoice) {
        '1' { return "Access" }
        '2' { return "OneDrive" }
        '3' { return "Outlook" }
        '4' { return "Word" }
        '5' { return "Excel" }
        '6' { return "OneNote" }
        '7' { return "Publisher" }
        '8' { return "PowerPoint" }
        '9' { return "Teams" }
        '10' { return "ProjectPro2019Volume" }
        '11' { return "ProjectPro2021Volume" }
        '12' { return "ProjectPro2024Volume" }
        '0' { return $null }
        default { Write-Host "Invalid selection, please choose between 0-12." -ForegroundColor Red; return Show-ComponentsMenu }
    }
}

function Install-Software {
    Write-Host "Installing software with Chocolatey..." -ForegroundColor Green

    # Install Google Chrome
    choco install googlechrome -y

    # Install Mozilla Firefox
    choco install firefox -y

    # Install WinRAR
    choco install winrar -y

    # Install VLC Player
    choco install vlc -y

    # Install Microsoft Office if needed
    $officeVersion = Show-OfficeVersionMenu
    if ($officeVersion) {
        $officeChannel = Show-ChannelMenu
        if ($officeChannel) {
            $officeComponents = Show-ComponentsMenu
            if ($officeComponents) {
                Write-Host "`nDownloading Office version: $officeVersion" -ForegroundColor Yellow
                Write-Host "Using channel: $officeChannel" -ForegroundColor Yellow
                Write-Host "Selected components: $officeComponents" -ForegroundColor Yellow
                
                # Insert your actual Office download logic here using the selected $officeVersion, $officeChannel, and $officeComponents
                Write-Host "Downloading and installing Office..." -ForegroundColor Yellow
            }
        }
    }
}

function Set-Up-Computer {
    Write-Host "Automating PC setup..." -ForegroundColor Green

    # Set hostname
    $hostname = Read-Host "Enter the hostname for this PC"
    Rename-Computer -NewName $hostname -Force

    # Set time zone
    $timeZone = Read-Host "Enter your time zone"
    Set-TimeZone -Id $timeZone

    # Install software (Chrome, Firefox, etc.)
    Install-Software
}

# Main Program Flow
Set-Up-Computer
