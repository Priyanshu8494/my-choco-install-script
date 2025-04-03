# PC SETUP - PowerShell Script
# Automates software installation and PC setup

# Function to check if Winget is installed and install it if missing
function Install-Winget {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "Winget not found. Installing..."
        try {
            Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile "$env:TEMP\Microsoft.DesktopAppInstaller.appxbundle"
            Add-AppxPackage -Path "$env:TEMP\Microsoft.DesktopAppInstaller.appxbundle"
            Write-Host "Winget installed successfully."
        } catch {
            Write-Host "Failed to install Winget: $_"
        }
    } else {
        Write-Host "Winget is already installed."
    }
}

# Function to install Office 2024
function Install-Office2024 {
    Write-Host "Installing Microsoft Office 2024..."
    $odtPath = "$env:TEMP\ODT"
    if (-not (Test-Path $odtPath)) {
        New-Item -ItemType Directory -Path $odtPath -Force | Out-Null
    }
    # Download Office Deployment Tool
    try {
        Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117" -OutFile "$env:TEMP\ODT.exe"
        Start-Process -FilePath "$env:TEMP\ODT.exe" -ArgumentList "/quiet /extract:$odtPath" -Wait
    } catch {
        Write-Host "Failed to download or extract Office Deployment Tool: $_"
        return
    }

    # Create Configuration File
    $xmlContent = @"
<Configuration>
<Add OfficeClientEdition="64" Channel="Current">
<Product ID="ProPlus2024Retail">
<Language ID="en-us"/>
</Product>
</Add>
<Property Name="AUTOACTIVATE" Value="1"/>
<Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
<Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@
    $xmlPath = "$odtPath\configuration.xml"
    $xmlContent | Set-Content -Path $xmlPath -Encoding UTF8

    # Run Office Installation
    try {
        Start-Process -FilePath "$odtPath\setup.exe" -ArgumentList "/configure $xmlPath" -Wait -NoNewWindow
        Write-Host "Office 2024 installation complete."
    } catch {
        Write-Host "Failed to install Office 2024: $_"
    }
}

# Main Menu Function
function Show-Menu {
    Write-Host "PC SETUP - Select an option:"
    Write-Host "1. Install Essential Software"
    Write-Host "2. Install Microsoft Office 2024"
    Write-Host "3. Run System Activation Toolkit"
    Write-Host "4. Update All Installed Software"
    $choice = Read-Host "Enter choice"
    switch ($choice) {
        "1" { 
            Write-Host "Installing essential software..."
            # Add commands to install essential software here
        }
        "2" { Install-Office2024 }
        "3" { 
            Write-Host "Running activation toolkit..."
            # Add commands to run activation toolkit here
        }
        "4" { 
            Write-Host "Updating installed software..."
            # Add commands to update all installed software here
        }
        default { Write-Host "Invalid option. Exiting..." }
    }
}

# Run the menu
Show-Menu
