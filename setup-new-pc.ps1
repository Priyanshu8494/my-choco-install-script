# setup.ps1 - Priyanshu Suryavanshi PC Setup Toolkit

function Show-Header {
    Clear-Host
    Write-Host "`n============================================================" -ForegroundColor White
    Write-Host "     Priyanshu Suryavanshi PC Setup Toolkit" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "============================================================`n" -ForegroundColor White
}

function Show-Menu {
    param (
        [string]$StatusMessage = "", 
        [string]$StatusColor = "Yellow"
    )
    
    Show-Header

    Write-Host "📌 Status:" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "   ➔ $StatusMessage" -ForegroundColor $StatusColor
    } else {
        Write-Host "   ➔ Ready to assist with your setup!" -ForegroundColor Green
    }

    Write-Host "`n🧪 Main Menu:" -ForegroundColor Yellow
    Write-Host "   [1] 📦 Install Essential Software" -ForegroundColor White
    Write-Host "   [2] 💼 Install MS Office Suite" -ForegroundColor White
    Write-Host "   [3] 🔑 System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "   [4] 🔄 Update All Installed Software (via Winget)" -ForegroundColor White
    Write-Host "   [0] ❌ Exit" -ForegroundColor Red

    Write-Host "`n💡 Tip: Use numbers to navigate (e.g., 1 for software install)"
    Write-Host "`n============================================================" -ForegroundColor DarkCyan
}

function Show-Loading {
    param([string]$Message = "Processing")
    $dots = ".", "..", "..."
    for ($i = 0; $i -lt 3; $i++) {
        Write-Host "`r$Message$dots[$i]" -NoNewline
        Start-Sleep -Milliseconds 300
    }
    Write-Host "`r$Message... Done!`n" -ForegroundColor Green
}

function Ensure-PackageManagers {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "❌ Winget is not installed or not working properly!" -ForegroundColor Red
        Write-Host "➡️ Please manually install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        Read-Host "Press any key to return to the menu..."
        exit
    }

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "🍫 Chocolatey is not installed. Installing now..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    }
}

function Install-NormalSoftware {
    # -- Existing unchanged --
    Write-Host "Installing Essential Software..." -ForegroundColor Yellow
}

function Install-MSOffice {
    Write-Host "`n📆 Full Office (ODT) - 64-bit | Channel: Monthly | Language: en-us" -ForegroundColor Cyan

    $officeEdition = "O365ProPlusRetail"
    $channel = "Monthly"
    $language = "en-us"

    $odtDownloadUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $workingDir = "$env:TEMP\OfficeODT"
    $setupExe = "$workingDir\setup.exe"
    $configPath = "$workingDir\config.xml"

    if (!(Test-Path $workingDir)) { New-Item -ItemType Directory -Path $workingDir -Force | Out-Null }

    Write-Host "`n[+] Downloading Office Deployment Tool..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $setupExe -UseBasicParsing -ErrorAction SilentlyContinue

    if (!(Test-Path $setupExe)) {
        Write-Host "[✘] Failed to download Office Deployment Tool. Check your internet or URL." -ForegroundColor Red
        Read-Host "Press any key to return to the menu..."
        return
    }

    Write-Host "[+] Extracting..." -ForegroundColor Yellow
    Start-Process -FilePath $setupExe -ArgumentList "/quiet /extract:$workingDir" -Wait

    $extractedSetup = Get-ChildItem -Path $workingDir -Recurse -Filter "setup.exe" | Select-Object -First 1

    if (!$extractedSetup) {
        Write-Host "[✘] setup.exe not found after extraction. Install failed." -ForegroundColor Red
        Read-Host "Press any key to return to the menu..."
        return
    }

    $setupExe = $extractedSetup.FullName

    Write-Host "[+] Generating config.xml..." -ForegroundColor Yellow
    $configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$officeEdition">
      <Language ID="$language" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
    $configXml | Set-Content -Path $configPath -Encoding UTF8

    Write-Host "[+] Installing Office silently..." -ForegroundColor Yellow
    Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configPath`"" -Wait

    Write-Host "✅ Office installation complete (check Start Menu)." -ForegroundColor Green
    Read-Host "Press any key to return to the menu..."
}

function Update-AllSoftware {
    Write-Host "Updating all software..." -ForegroundColor Yellow
    # -- Placeholder --
}

function Invoke-Activation {
    Write-Host "Running Activation Toolkit..." -ForegroundColor Yellow
    Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
    Read-Host "Press any key to return to the menu..."
}

# MAIN
Ensure-PackageManagers

Do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '0' {
            Write-Host "`n👋 Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            break
        }
        default {
            Show-Menu -StatusMessage "⚠️ Invalid selection! Please choose between 0-4." -StatusColor "Red"
        }
    }
} While ($true)
