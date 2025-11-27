<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
.NOTES
  - Clean UI Edition
  - Work in progress.
#>

# -------------------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------------------
function Show-Header {
    Clear-Host
    $width = 70
    $line = "=" * $width
    
    Write-Host $line -ForegroundColor Cyan
    Write-Host "    ____  ______   _____ ______ ______  __  __  ____" -ForegroundColor Magenta
    Write-Host "   / __ \/ ____/  / ___// ____//_  __/ / / / / / __ \" -ForegroundColor Magenta
    Write-Host "  / /_/ / /       \__ \/ __/    / /   / / / / / /_/ /" -ForegroundColor Cyan
    Write-Host " / ____/ /___    ___/ / /___   / /   / /_/ / / ____/" -ForegroundColor Cyan
    Write-Host "/_/    \____/   /____/_____/  /_/    \____/ /_/     " -ForegroundColor Blue
    Write-Host ""
    Write-Host "      Priyanshu Suryavanshi PC Setup Toolkit" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host "      Automated Setup & Activation Utility" -ForegroundColor Gray
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    param (
        [string]$StatusMessage = "",
        [string]$StatusColor = "Yellow"
    )
    Show-Header
    
    # Status Section
    Write-Host " [ SYSTEM STATUS ]" -ForegroundColor Cyan
    if ($StatusMessage) {
        Write-Host "  >> $StatusMessage" -ForegroundColor $StatusColor
    }
    else {
        Write-Host "  >> System Ready. Waiting for input..." -ForegroundColor Green
    }
    Write-Host ""

    # Menu Section
    Write-Host " [ MAIN MENU ]" -ForegroundColor Yellow
    
    $menuItems = @(
        @{ Key = "1"; Label = "Install Essential Software"; Desc = "Browsers, Media, Utilities" },
        @{ Key = "2"; Label = "Install MS Office Suite"; Desc = "Office 2021 Pro Plus" },
        @{ Key = "3"; Label = "System Activation Toolkit"; Desc = "Windows & Office Activation" },
        @{ Key = "4"; Label = "Update All Software"; Desc = "Upgrade via Winget" },
        @{ Key = "5"; Label = "Advanced Toolkit"; Desc = "WinUtil by Chris Titus" },
        @{ Key = "0"; Label = "Exit Application"; Desc = "Close the script" }
    )

    foreach ($item in $menuItems) {
        Write-Host "  [" -NoNewline -ForegroundColor DarkGray
        Write-Host "$($item.Key)" -NoNewline -ForegroundColor Cyan
        Write-Host "] " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($item.Label)" -NoNewline -ForegroundColor White
        if ($item.Desc) {
            Write-Host " - $($item.Desc)" -ForegroundColor DarkGray
        }
        else {
            Write-Host ""
        }
    }
    
    Write-Host ""
    Write-Host "  Tip: Enter the number corresponding to your choice." -ForegroundColor DarkCyan
    Write-Host " ======================================================================" -ForegroundColor Cyan
}

function Show-Loading {
    param([string]$Message = "Processing")
    $spinner = @('|', '/', '-', '\')
    Write-Host -NoNewline "  ⏳ $Message " -ForegroundColor Yellow
    
    for ($i = 0; $i -lt 15; $i++) {
        foreach ($char in $spinner) {
            Write-Host -NoNewline "`b$char" -ForegroundColor Cyan
            Start-Sleep -Milliseconds 100
        }
    }
    Write-Host "`b✅ Done!" -ForegroundColor Green
    Write-Host ""
}

# -------------------------------------------------------------------------
# Ensure package managers present
# -------------------------------------------------------------------------
function Ensure-PackageManagers {
    # Check winget
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "`n❌ Winget is not installed or not available in PATH!" -ForegroundColor Red
        Write-Host "➡️ Please install Winget from: https://aka.ms/getwinget" -ForegroundColor Yellow
        # Do not automatically exit — allow user to continue for tasks that don't need winget
    }
    else {
        Write-Host "`n✅ Winget detected." -ForegroundColor Green
    }

    # Check Chocolatey, install if missing
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "🍫 Chocolatey not found. Installing now..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            $chocoInstallScript = 'https://community.chocolatey.org/install.ps1'
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString($chocoInstallScript))
            Write-Host "✅ Chocolatey installation attempted. You may need to re-open shell." -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Chocolatey install failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "✅ Chocolatey detected." -ForegroundColor Green
    }
}

# -------------------------------------------------------------------------
# Utility: Create Desktop Shortcut
# -------------------------------------------------------------------------
function New-DesktopShortcut {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [Parameter(Mandatory)][string]$ShortcutName,
        [string]$Arguments = ""
    )
    try {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$desktop\$ShortcutName.lnk")
        $Shortcut.TargetPath = $TargetPath
        if ($Arguments) { $Shortcut.Arguments = $Arguments }
        $Shortcut.IconLocation = "$TargetPath,0"
        $Shortcut.Save()
        return $true
    }
    catch {
        return $false
    }
}

# -------------------------------------------------------------------------
# Software installation
# -------------------------------------------------------------------------
function Install-AnyDeskDirectly {
    $anydeskUrl = "https://download.anydesk.com/AnyDesk.exe"
    $installPath = "C:\AnyDesk"
    $exePath = Join-Path $installPath "AnyDesk.exe"

    Write-Host "`n [ INSTALLING ANYDESK ]" -ForegroundColor Cyan
    Write-Host " Target: $installPath" -ForegroundColor Gray
    try {
        if (-not (Test-Path $installPath)) {
            New-Item -ItemType Directory -Path $installPath -Force | Out-Null
        }
        Write-Host "⏬ Downloading AnyDesk..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $anydeskUrl -OutFile $exePath -UseBasicParsing -ErrorAction Stop

        Write-Host "🚀 Running AnyDesk installer (silent flags may vary)..." -ForegroundColor Yellow
        Start-Process -FilePath $exePath -ArgumentList "--install" -Wait -NoNewWindow -ErrorAction SilentlyContinue

        # Create shortcut (if executable exists)
        if (Test-Path $exePath) {
            if (New-DesktopShortcut -TargetPath $exePath -ShortcutName "AnyDesk") {
                Write-Host "✅ AnyDesk installed and shortcut created." -ForegroundColor Green
            }
            else {
                Write-Host "✅ AnyDesk downloaded. Shortcut creation failed or skipped." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "❌ AnyDesk installer file not found after download." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "❌ Error installing AnyDesk: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-UltraViewerDirectly {
    $uvUrl = "https://ultraviewer.net/UltraViewer_setup.exe"
    $installPath = "C:\UltraViewer"
    $exePath = Join-Path $installPath "UltraViewer_setup.exe"

    Write-Host "`n [ INSTALLING ULTRAVIEWER ]" -ForegroundColor Cyan
    Write-Host " Target: $installPath" -ForegroundColor Gray
    try {
        if (-not (Test-Path $installPath)) {
            New-Item -ItemType Directory -Path $installPath -Force | Out-Null
        }
        Write-Host "⏬ Downloading UltraViewer..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $uvUrl -OutFile $exePath -UseBasicParsing -ErrorAction Stop

        Write-Host "🚀 Running UltraViewer installer silently..." -ForegroundColor Yellow
        Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT", "/NORESTART" -Wait -NoNewWindow -ErrorAction SilentlyContinue

        $installedPath = "C:\Program Files\UltraViewer\UltraViewer.exe"
        if (Test-Path $installedPath) {
            if (New-DesktopShortcut -TargetPath $installedPath -ShortcutName "UltraViewer") {
                Write-Host "✅ UltraViewer installed and shortcut created." -ForegroundColor Green
            }
            else {
                Write-Host "✅ UltraViewer installed. Shortcut creation failed or skipped." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "⚠️ UltraViewer installer ran but expected EXE not found; verify install location." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "❌ Error installing UltraViewer: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-NormalSoftware {
    Ensure-PackageManagers

    $softwareList = @(
        @{Name = "Google Chrome"; ID = "Google.Chrome" },
        @{Name = "Mozilla Firefox"; ID = "Mozilla.Firefox" },
        @{Name = "WinRAR"; ID = "RARLab.WinRAR" },
        @{Name = "VLC Player"; ID = "VideoLAN.VLC" },
        @{Name = "PDF Reader (Sumatra)"; ID = "SumatraPDF.SumatraPDF" },
        @{Name = "AnyDesk"; ID = "custom-anydesk" },
        @{Name = "UltraViewer"; ID = "custom-ultraviewer" }
    )

    Write-Host "`n [ SOFTWARE SELECTION ]" -ForegroundColor Yellow
    Write-Host " Select software to install (e.g., 1,3,5 or 'all'):" -ForegroundColor Gray
    Write-Host ""
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        $num = $i + 1
        $name = $softwareList[$i].Name
        Write-Host "  [" -NoNewline -ForegroundColor DarkGray
        Write-Host "$num" -NoNewline -ForegroundColor Cyan
        Write-Host "] $name" -ForegroundColor White
    }

    $selection = Read-Host "Enter selection (e.g., 1,3,5). Or 'all' to install everything"
    if ($selection.Trim().ToLower() -eq 'all') {
        $selectedIndices = 1..$softwareList.Count
    }
    else {
        $selectedIndices = @()
        foreach ($token in ($selection -split ",")) {
            $t = $token.Trim()
            if ($t -match '^\d+$') {
                $selectedIndices += [int]$t
            }
        }
    }

    foreach ($index in $selectedIndices | Sort-Object -Unique) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            if ($app.ID -eq "custom-anydesk") {
                Install-AnyDeskDirectly
            }
            elseif ($app.ID -eq "custom-ultraviewer") {
                Install-UltraViewerDirectly
            }
            else {
                Write-Host "`n📦 Installing $($app.Name) via winget..." -ForegroundColor Gray
                try {
                    if (Get-Command winget -ErrorAction SilentlyContinue) {
                        # Use --accept flags to avoid prompts
                        Start-Process -FilePath "winget" -ArgumentList "install --id $($app.ID) --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -ErrorAction Stop
                        Write-Host "✅ $($app.Name) installed successfully!" -ForegroundColor Green
                    }
                    else {
                        Write-Host "⚠️ winget not available. Skipping $($app.Name)." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "❌ Failed to install $($app.Name): $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "  ⚠️ Invalid selection: $index" -ForegroundColor Red
        }
    }

    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Office installation
# -------------------------------------------------------------------------
function Install-OfficeOnline {
    # --- CONFIGURATION ---
    # Updated Link (1.33.zip)
    $ZipUrl = "https://github.com/Priyanshu8494/my-choco-install-script/raw/refs/heads/main/Office%20Installer%201.33.zip"

    $BaseDir = "C:\Temp\Office_Auto_Install"
    $ZipPath = "$BaseDir\OfficeInstaller.zip"

    # Clear Screen
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      OFFICE INSTALLER (UPDATED LINK)     " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan

    # --- STEP 0: PREPARE FOLDER & BYPASS DEFENDER ---
    Write-Host "[1/5] Preparing Security Exclusions..." -ForegroundColor Yellow

    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force | Out-Null }

    try {
        # Defender exclusion to prevent antivirus from deleting the setup
        Add-MpPreference -ExclusionPath $BaseDir -ErrorAction SilentlyContinue
        Write-Host "      Defender Exclusion Added." -ForegroundColor Green
    }
    catch {
        Write-Host "      Warning: Admin rights needed for exclusion." -ForegroundColor Red
    }

    # --- STEP 1: DOWNLOAD ---
    Write-Host "[2/5] Downloading Installer..." -ForegroundColor Yellow
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipPath -ErrorAction Stop
        Write-Host "      Download Complete!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Download Failed! Check your internet connection." -ForegroundColor Red
        Write-Host "Details: $_" -ForegroundColor DarkGray
        return
    }

    # --- STEP 2: UNBLOCK & EXTRACT ---
    Write-Host "[3/5] Unblocking & Extracting..." -ForegroundColor Yellow

    # SmartScreen Bypass
    Unblock-File -Path $ZipPath

    # Extract
    Expand-Archive -Path $ZipPath -DestinationPath $BaseDir -Force
    Write-Host "      Extraction Complete!" -ForegroundColor Green

    # --- STEP 3: IDENTIFY FILE ---
    Write-Host "[4/5] Selecting Installer..." -ForegroundColor Yellow

    $Installer64 = Get-ChildItem -Path $BaseDir -Filter "Office Installer.exe" -Recurse | Select-Object -First 1
    $Installer32 = Get-ChildItem -Path $BaseDir -Filter "Office Installer x86.exe" -Recurse | Select-Object -First 1

    $TargetFile = $null
    if ([Environment]::Is64BitOperatingSystem -and $Installer64) {
        $TargetFile = $Installer64
    }
    elseif ($Installer32) {
        $TargetFile = $Installer32
    }

    # --- STEP 4: RUN ---
    if ($TargetFile) {
        Write-Host "[5/5] Launching Installer..." -ForegroundColor Cyan
        
        # --- UPDATED ENGLISH MESSAGE ---
        Write-Host "      NOTE: A window will open. Click 'Install' to proceed." -ForegroundColor Magenta
        
        Start-Process -FilePath $TargetFile.FullName -WorkingDirectory $TargetFile.Directory.FullName -Wait
        
        Write-Host "      Process Finished." -ForegroundColor Green
    }
    else {
        Write-Host "Error: 'Office Installer.exe' not found in zip!" -ForegroundColor Red
    }

    # --- CLEANUP ---
    if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
}


function Create-OfficeOfflineInstaller {
    $targetDir = "C:\OfficeOfflineInstaller"
    $setupExe = Join-Path $targetDir "setup.exe"
    $configFile = Join-Path $targetDir "Configuration.xml"

    Write-Host "`n [ OFFICE OFFLINE INSTALLER CREATOR ]" -ForegroundColor Yellow
    Write-Host " This will download Office files for offline use." -ForegroundColor Gray
    Write-Host " Target Directory: $targetDir" -ForegroundColor Cyan
    
    $userPath = Read-Host " Press Enter to use default path, or type a new path"
    if (-not [string]::IsNullOrWhiteSpace($userPath)) {
        $targetDir = $userPath
        $setupExe = Join-Path $targetDir "setup.exe"
        $configFile = Join-Path $targetDir "Configuration.xml"
    }

    try {
        # 1. Create Directory
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        # 2. Download ODT Setup.exe
        Write-Host "⏬ Downloading Office Deployment Tool..." -ForegroundColor Yellow
        $odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
        Invoke-WebRequest -Uri $odtUrl -OutFile $setupExe -UseBasicParsing -ErrorAction Stop

        # 3. Create Configuration.xml
        Write-Host "📝 Generating configuration file..." -ForegroundColor Yellow
        $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
        $xmlContent | Set-Content -Path $configFile -Encoding UTF8

        # 4. Run Installer in Download Mode
        Write-Host "⏬ Starting Office Download..." -ForegroundColor Yellow
        Write-Host "   (This will take a while depending on internet speed. Please wait...)" -ForegroundColor Gray
        
        Start-Process -FilePath $setupExe -ArgumentList "/download Configuration.xml" -WorkingDirectory $targetDir -Wait -NoNewWindow -ErrorAction Stop
        
        Write-Host "✅ Office offline files downloaded successfully to: $targetDir" -ForegroundColor Green
        Write-Host "   To install later, run: setup.exe /configure Configuration.xml" -ForegroundColor Gray
    }
    catch {
        Write-Host "❌ Office download failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Install-MSOffice {
    Write-Host "`n [ OFFICE INSTALLATION ]" -ForegroundColor Yellow
    Write-Host " Choose installation method:" -ForegroundColor Gray
    Write-Host "   [1] Office Online Installer" -ForegroundColor White
    Write-Host "   [2] Office Offline Installer" -ForegroundColor White
    Write-Host "   [0] Go Back" -ForegroundColor DarkGray
    $subChoice = Read-Host "Enter your choice [0-2]"

    switch ($subChoice) {
        '1' {
            Install-OfficeOnline
            Read-Host "Press Enter to return to the menu..."
        }
        '2' {
            Create-OfficeOfflineInstaller
            Read-Host "Press Enter to return to the menu..."
        }
        default {
            return
        }
    }
}

# -------------------------------------------------------------------------
# Update all software
# -------------------------------------------------------------------------
function Update-AllSoftware {
    Ensure-PackageManagers
    Write-Host "`n [ SYSTEM UPDATE ]" -ForegroundColor Yellow
    Write-Host " Checking for software updates via winget..." -ForegroundColor Gray
    try {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Start-Process -FilePath "winget" -ArgumentList "upgrade --all --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -ErrorAction Stop
            Write-Host "✅ All installed software updated successfully (winget)." -ForegroundColor Green
        }
        else {
            Write-Host "❌ winget not available. Cannot perform updates." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "❌ Failed to update some software: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Activation (remote)
# -------------------------------------------------------------------------
function Invoke-Activation {
    Write-Host "`n [ SYSTEM ACTIVATION ]" -ForegroundColor Yellow
    Write-Host " Running System Activation Toolkit (remote)... " -ForegroundColor Gray
    Write-Host "⚠️ This will execute a remote script. Ensure you trust the source before proceeding." -ForegroundColor Magenta

    try {
        # NOTE: Remote execution is potentially dangerous. Keep it as-is per original but wrapped.
        irm https://get.activated.win | iex
    }
    catch {
        Write-Host "❌ Activation script failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Advanced Toolkit (remote)
# -------------------------------------------------------------------------
function Invoke-AdvancedToolkit {
    Write-Host "`n [ ADVANCED TOOLKIT ]" -ForegroundColor Yellow
    Write-Host " Running Advanced Toolkit (remote script)..." -ForegroundColor Gray
    Write-Host "⚠️ This will execute a remote script (irm https://christitus.com/win | iex)." -ForegroundColor Magenta

    try {
        irm https://christitus.com/win | iex
        Write-Host "✅ Advanced Toolkit execution finished." -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Advanced Toolkit failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the menu..."
}

# -------------------------------------------------------------------------
# Main program
# -------------------------------------------------------------------------
# Ensure package managers checked at start (non-fatal)
Ensure-PackageManagers

do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-5]"
    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '5' { Invoke-AdvancedToolkit }
        '0' {
            Write-Host "`n👋 Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            break
        }
        default {
            Show-Menu -StatusMessage "⚠️ Invalid selection! Please choose between 0-5." -StatusColor "Red"
        }
    }
} while ($true)
