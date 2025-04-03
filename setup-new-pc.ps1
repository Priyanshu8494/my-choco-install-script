# Set Console Encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Main Menu
function Show-Menu {
    param (
        [string]$Title = "PC SETUP - Main Menu"
    )
    Clear-Host
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "           $Title            " -ForegroundColor Yellow
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "1: Install Essential Software" -ForegroundColor Green
    Write-Host "2: Install MS Office" -ForegroundColor Green
    Write-Host "3: Run System Activation Toolkit" -ForegroundColor Green
    Write-Host "4: Update All Installed Software" -ForegroundColor Green
    Write-Host "0: Exit" -ForegroundColor Red
    Write-Host "==================================" -ForegroundColor Cyan
}

# Progress Bar Example Function
function Show-Progress {
    param (
        [string]$Activity = "Processing",
        [int]$PercentComplete = 0
    )
    Write-Progress -Activity $Activity -Status "Please wait..." -PercentComplete $PercentComplete
}

# Function to Install Essential Software
function Install-EssentialSoftware {
    Write-Host "Installing Essential Software..." -ForegroundColor Cyan
    Show-Progress -Activity "Installing Software" -PercentComplete 10
    winget install --id=Google.Chrome --silent --accept-source-agreements --accept-package-agreements
    winget install --id=Mozilla.Firefox --silent --accept-source-agreements --accept-package-agreements
    winget install --id=RARLab.WinRAR --silent --accept-source-agreements --accept-package-agreements
    winget install --id=VideoLAN.VLC --silent --accept-source-agreements --accept-package-agreements
    Show-Progress -Activity "Installing Software" -PercentComplete 100
    Write-Host "Installation Complete!" -ForegroundColor Green
}

# Function to Install MS Office
function Install-MSOffice {
    Write-Host "Downloading and Installing MS Office..." -ForegroundColor Cyan
    Invoke-Expression "& { $(Invoke-RestMethod 'https://raw.githubusercontent.com/example/MSOfficeInstaller/main/install.ps1') }"
    Write-Host "MS Office Installation Complete!" -ForegroundColor Green
}

# Function to Run Activation Toolkit
function Run-ActivationToolkit {
    Write-Host "Downloading Activation Toolkit..." -ForegroundColor Cyan
    $activationPath = "$env:TEMP\activate.exe"
    Invoke-WebRequest -Uri "https://activationtoolkit.example.com/activate.exe" -OutFile $activationPath
    Write-Host "Running Activation Toolkit..." -ForegroundColor Cyan
    Start-Process -FilePath $activationPath -Wait
    Write-Host "Activation Complete!" -ForegroundColor Green
}

# Function to Update Software
function Update-Software {
    Write-Host "Updating Installed Software..." -ForegroundColor Cyan
    winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
    Write-Host "All Software Updated Successfully!" -ForegroundColor Green
}

# Main Menu Execution
Do {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    Switch ($choice) {
        "1" { Install-EssentialSoftware }
        "2" { Install-MSOffice }
        "3" { Run-ActivationToolkit }
        "4" { Update-Software }
        "0" { Write-Host "Exiting..." -ForegroundColor Red; Exit }
        Default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
    }
    Pause
} While ($true)
