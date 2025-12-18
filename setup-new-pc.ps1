# .SYNOPSIS
#   Priyanshu Suryavanshi PC Setup Toolkit
# .DESCRIPTION
#   Automated PC setup with software installation and system activation
# .NOTES
#   - Clean UI Edition
#   - Work in progress.

# Force TLS 1.2 for all web requests (Fixes download errors)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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
        @{ Key = "6"; Label = "Ram Optimization"; Desc = "Global Ram Optimization Script" },
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
# Global RAM Optimizer (Embedded)
# -------------------------------------------------------------------------
function Launch-GlobalOptimizer {
    Write-Host "`n [ RAM OPTIMIZATION ]" -ForegroundColor Yellow
    Write-Host " Extracting Global Ram Optimization Suite..." -ForegroundColor Gray
    
    # Base64 Encoded Content of 'Global Ram Optimization.ps1'
    $b64 = "PCMNCi5TWU5PUFNJUw0KICAgIEdsb2JhbCAiU21hcnQiIFJBTSBPcHRpbWl6YXRpb24gTWFuYWdlbWVudCBTY3JpcHQuDQogICAgDQouREVTQ1JJUFRJT04NCiAgICBQcm92aWRlcyBhIG1lbnUgdG86DQogICAgMS4gUnVuIFNtYXJ0IEdsb2JhbCBPcHRpbWl6YXRpb24gKFRlc3QgTW9kZSAtIFZpc2libGUpLg0KICAgIDIuIEluc3RhbGwgUGVybWFuZW50IFNtYXJ0IE9wdGltaXphdGlvbiAoQmFja2dyb3VuZCBNb2RlIC0gU2lsZW50KS4NCiAgICAzLiBSZW1vdmUgR2xvYmFsIE9wdGltaXphdGlvbi4NCiAgICANCiAgICBGZWF0dXJlczoNCiAgICAtIFNjYW5zIGZvciBBTlkgcHJvY2VzcyB1c2luZyA+IDEwMCBNQiBSQU0uDQogICAgLSBTYWZlbHkgZXhjdWRlcyBjcml0aWNhbCBXaW5kb3dzIFN5c3RlbSBwcm9jZXNzZXMuDQogICAgLSBUcmltcyBtZW1vcnkgdXNpbmcgRW1wdHlXb3JraW5nU2V0IEFQSS4NCiAgICANCi5OT1RFUw0KICAgIEZpbGUgTmFtZTogR2xvYmFsIFJhbSBPcHRpbWl6YXRpb24ucHMxDQogICAgTXVzdCBiZSBydW4gYXMgQWRtaW5pc3RyYXRvci4NCiM+DQoNCiMgLS0tIEFkbWluIENoZWNrIC0tLQ0KJGN1cnJlbnRQcmluY2lwYWwgPSBOZXctT2JqZWN0IFNlY3VyaXR5LlByaW5jaXBhbC5XaW5kb3dzUHJpbmNpcGFsKFtTZWN1cml0eS5QcmluY2lwYWwuV2luZG93c0lkZW50aXR5XTo6R2V0Q3VycmVudCgpKQ0KaWYgKC1ub3QgJGN1cnJlbnRQcmluY2lwYWwuSXNJblJvbGUoW1NlY3VyaXR5LlByaW5jaXBhbC5XaW5kb3dzQnVpbHRJblJvbGVdOjpBZG1pbmlzdHJhdG9yKSkgew0KICAgICMgQWRkZWQgLU5vRXhpdCBzbyB0aGUgd2luZG93IHN0YXlzIG9wZW4gb24gV2luMTEgZXZlbiBpZiB0aGVyZSBpcyBhbiBlcnJvcg0KICAgIFN0YXJ0LVByb2Nlc3MgcG93ZXJzaGVsbC5leGUgLUFyZ3VtZW50TGlzdCAiLU5vUHJvZmlsZSAtTm9FeGl0IC1FeGVjdXRpb25Qb2xpY3kgQnlwYXNzIC1GaWxlIGAiJFBTQ29tbWFuZFBhdGhgIiIgLVZlcmIgUnVuQXMNCiAgICBFeGl0DQp9DQoNCiMgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09DQojICAgICAgRU1CRURERUQgU0NSSVBUUyAoU09VUkNFIENPREUpDQojID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQ0KDQojIDEuIFNNQVJUIEdMT0JBTCBPUFRJTUlaRVIgKFRFU1QgTU9ERSAtIFZJU0lCTEUpDQokU2NyaXB0X1Rlc3RfR2xvYmFsID0gQCcNCiMgLS0tIEF1dG8tR2VuZXJhdGVkIFNtYXJ0IEdsb2JhbCBSQU0gT3B0aW1pemVyIC0tLQ0KaWYgKC1ub3QgKCJNZW1vcnlUcmltbWVyIiAtYXMgW3R5cGVdKSkgew0KICAgICRjb2RlID0gQCINCnVzaW5nIFN5c3RlbTsNCnVzaW5nIFN5c3RlbS5SdW50aW1lLkludGVyb3BTZXJ2aWNlczsNCnVzaW5nIFN5c3RlbS5EaWFnbm9zdGljczsNCnB1YmxpYyBjbGFzcyBNZW1vcnlUcmltbWVyIHsNCiAgICBbRGxsSW1wb3J0KCJwc2FwaS5kbGwiKV0NCiAgICBwdWJsaWMgc3RhdGljIGV4dGVybiBib29sIEVtcHR5V29ya2luZ1NldChJbnRQdHIgaFByb2Nlc3MpOw0KICAgIHB1YmxpYyBzdGF0aWMgdm9pZCBUcmltUHJvY2VzcyhpbnQgcGlkKSB7DQogICAgICAgIHRyeSB7DQogICAgICAgICAgICBQcm9jZXNzIHAgPSBQcm9jZXNzLkdldFByb2Nlc3NCeUlkKHBpZCk7DQogICAgICAgICAgICBFbXB0eVdvcmtpbmdTZXQocC5IYW5kbGUpOw0KICAgICAgICB9IGNhdGNoIHsgfQ0KICAgIH0NCn0NCiJADQogICAgQWRkLVR5cGUgLVR5cGVEZWZpbml0aW9uICRjb2RlDQp9DQoNCiRFeGNsdXNpb25zID0gQCgNCiAgICAiSWRsZSIsICJTeXN0ZW0iLCAiUmVnaXN0cnkiLCAic21zcyIsICJjc3JzcyIsICJ3aW5pbml0IiwgInNlcnZpY2VzIiwgImxzYXNzIiwgDQogICAgIndpbmxvZ29uIiwgImZvbnRkcnZob3N0IiwgImR3bSIsICJNZW1vcnkgQ29tcHJlc3Npb24iLCAiTXNNcEVuZyIsICJ0YXNrbWdyIg0KKQ0KDQpXcml0ZS1Ib3N0ICI9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0iIC1Gb3JlZ3JvdW5kQ29sb3IgQ3lhbg0KV3JpdGUtSG9zdCAiICAgR0xPQkFMIFNNQVJUIFJBTSBPUFRJTUlaRVIgKFRFU1RJTkcpICAgIiAtRm9yZWdyb3VuZENvbG9yIEN5YW4NCldyaXRlLUhvc3QgIj09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSIgLUZvcmVncm91bmRDb2xvciBDeWFuDQpXcml0ZS1Ib3N0ICJUYXJnZXQ6IFByb2Nlc3NlcyA+IDEwMCBNQiAoVGVhbXMsIEJyb3dzZXJzLCBldGMuKSIgLUZvcmVncm91bmRDb2xvciBHcmVlbg0KV3JpdGUtSG9zdCAiQWN0aW9uOiBUcmltIFdvcmtpbmcgU2V0IiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuDQpXcml0ZS1Ib3N0ICJFeGNsdWRpbmc6IFdpbmRvd3MgU3lzdGVtIFByb2Nlc3NlcyIgLUZvcmVncm91bmRDb2xvciBHcmF5DQpXcml0ZS1Ib3N0ICJQcmVzcyBDdHJsK0MgdG8gc3RvcC4iIC1Gb3JlZ3JvdW5kQ29sb3IgWWVsbG93DQpXcml0ZS1Ib3N0ICItLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0iIC1Gb3JlZ3JvdW5kQ29sb3IgR3JheQ0KDQp0cnkgew0KICAgIHdoaWxlICgkdHJ1ZSkgew0KICAgICAgICAkdGltZXN0YW1wID0gR2V0LURhdGUgLUZvcm1hdCAiSEg6bW06c3MiDQogICAgICAgICR0b3RhbEZyZWVkQ3ljbGUgPSAwDQogICAgICAgICR0cmltbWVkQXBwcyA9IEAoKQ0KDQogICAgICAgICMgRmluZCBIaWdoIE1lbW9yeSBQcm9jZXNzZXMNCiAgICAgICAgJHRhcmdldHMgPSBHZXQtUHJvY2VzcyAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZSB8IFdoZXJlLU9iamVjdCB7IA0KICAgICAgICAgICAgJF8uV29ya2luZ1NldCAtZ3QgMTAwTUIgLWFuZCANCiAgICAgICAgICAgICRfLlByb2Nlc3NOYW1lIC1ub3RpbiAkRXhjbHVzaW9ucyANCiAgICAgICAgfQ0KDQogICAgICAgIGlmICgkdGFyZ2V0cykgew0KICAgICAgICAgICAgZm9yZWFjaCAoJHByb2MgaW4gJHRhcmdldHMpIHsNCiAgICAgICAgICAgICAgICAkYmVmb3JlID0gJHByb2MuV29ya2luZ1NldA0KICAgICAgICAgICAgICAgIA0KICAgICAgICAgICAgICAgICMgVFJJTQ0KICAgICAgICAgICAgICAgIFtNZW1vcnlUcmltbWVyXTo6VHJpbVByb2Nlc3MoJHByb2MuSWQpDQogICAgICAgICAgICAgICAgDQogICAgICAgICAgICAgICAgIyBNZWFzdXJlIFNhdmluZ3MNCiAgICAgICAgICAgICAgICB0cnkgeyAkcHJvYy5SZWZyZXNoKCk7ICRhZnRlciA9ICRwcm9jLldvcmtpbmdTZXQgfSBjYXRjaCB7ICRhZnRlciA9ICRiZWZvcmUgfQ0KICAgICAgICAgICAgICAgIA0KICAgICAgICAgICAgICAgICRzYXZlZCA9ICgkYmVmb3JlIC0gJGFmdGVyKSAvIDFNQg0KICAgICAgICAgICAgICAgIGlmICgkc2F2ZWQgLWd0IDEwKSB7DQogICAgICAgICAgICAgICAgICAgICR0b3RhbEZyZWVkQ3ljbGUgKz0gJHNhdmVkDQogICAgICAgICAgICAgICAgICAgICR0cmltbWVkQXBwcyArPSAiJCgkcHJvYy5Qcm9jZXNzTmFtZSkgKC0kKFttYXRoXTo6Um91bmQoJHNhdmVkLDApKU1CKSINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgIH0NCg0KICAgICAgICBpZiAoJHRvdGFsRnJlZWRDeWNsZSAtZ3QgMCkgew0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiWyR0aW1lc3RhbXBdIEZyZWVkICQoW21hdGhdOjpSb3VuZCgkdG90YWxGcmVlZEN5Y2xlLCAwKSkgTUIiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4gLU5vTmV3bGluZQ0KICAgICAgICAgICAgaWYgKCR0cmltbWVkQXBwcy5Db3VudCAtZ3QgMCkgew0KICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgfCAkKCR0cmltbWVkQXBwcyAtam9pbiAnLCAnKSIgLUZvcmVncm91bmRDb2xvciBEYXJrR3JheQ0KICAgICAgICAgICAgfSBlbHNlIHsgV3JpdGUtSG9zdCAiIiB9DQogICAgICAgIH0NCiAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgIyBIZWFydGJlYXQgdG8gc2hvdyBpdCdzIG5vdCBzdHVjaw0KICAgICAgICAgICAgIFdyaXRlLUhvc3QgIi4iIC1Ob05ld2xpbmUgLUZvcmVncm91bmRDb2xvciBEYXJrR3JheQ0KICAgICAgICB9DQogICAgICAgIA0KICAgICAgICBTdGFydC1TbGVlcCAtU2Vjb25kcyAzDQogICAgfQ0KfQ0KY2F0Y2ggeyANCiAgICBXcml0ZS1Ib3N0ICJgbkVycm9yIGluIExvb3A6ICRfIiAtRm9yZWdyb3VuZENvbG9yIFJlZA0KICAgIFN0YXJ0LVNsZWVwIC1TZWNvbmRzIDUNCn0NCldyaXRlLUhvc3QgImBuUHJlc3MgRW50ZXIgdG8gZXhpdC4uLiIgLUZvcmVncm91bmRDb2xvciBZZWxsb3cNClJlYWQtSG9zdA0KJ0ANCg0KIyAyLiBTTUFSVCBHTE9CQUwgT1BUSU1JWkVSIChQRVJNQU5FTlQgLSBTSUxFTlQpDQokU2NyaXB0X1Blcm1fR2xvYmFsID0gQCcNCiMgR2xvYmFsIFNtYXJ0IE9wdGltaXplciAoU2lsZW50KQ0KaWYgKC1ub3QgKCJNZW1vcnlUcmltbWVyIiAtYXMgW3R5cGVdKSkgew0KICAgICRjb2RlID0gQCINCnVzaW5nIFN5c3RlbTsNCnVzaW5nIFN5c3RlbS5SdW50aW1lLkludGVyb3BTZXJ2aWNlczsNCnVzaW5nIFN5c3RlbS5EaWFnbm9zdGljczsNCnB1YmxpYyBjbGFzcyBNZW1vcnlUcmltbWVyIHsNCiAgICBbRGxsSW1wb3J0KCJwc2FwaS5kbGwiKV0NCiAgICBwdWJsaWMgc3RhdGljIGV4dGVybiBib29sIEVtcHR5V29ya2luZ1NldChJbnRQdHIgaFByb2Nlc3MpOw0KICAgIHB1YmxpYyBzdGF0aWMgdm9pZCBUcmltUHJvY2VzcyhpbnQgcGlkKSB7DQogICAgICAgIHRyeSB7DQogICAgICAgICAgICBQcm9jZXNzIHAgPSBQcm9jZXNzLkdldFByb2Nlc3NCeUlkKHBpZCk7DQogICAgICAgICAgICBFbXB0eVdvcmtpbmdTZXQocC5IYW5kbGUpOw0KICAgICAgICB9IGNhdGNoIHsgfQ0KICAgIH0NCn0NCiJADQogICAgQWRkLVR5cGUgLVR5cGVEZWZpbml0aW9uICRjb2RlDQp9DQoNCiMgU2VsZi1IaWRpbmcNCiR3Q29kZSA9IEAiDQp1c2luZyBTeXN0ZW07IHVzaW5nIFN5c3RlbS5SdW50aW1lLkludGVyb3BTZXJ2aWNlczsNCnB1YmxpYyBjbGFzcyBXaW5kb3dNYW5hZ2VyIHsNCiAgICBbRGxsSW1wb3J0KCJrZXJuZWwzMi5kbGwiKV0gcHVibGljIHN0YXRpYyBleHRlcm4gSW50UHRyIEdldENvbnNvbGVXaW5kb3coKTsNCiAgICBbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIildIHB1YmxpYyBzdGF0aWMgZXh0ZXJuIGJvb2wgU2hvd1dpbmRvdyhJbnRQdHIgaFduZCwgaW50IG5DbWRTaG93KTsNCn0NCiJADQppZiAoLW5vdCAoIldpbmRvd01hbmFnZXIiIC1hcyBbdHlwZV0pKSB7IEFkZC1UeXBlIC1UeXBlRGVmaW5pdGlvbiAkd0NvZGUgfQ0KW1dpbmRvd01hbmFnZXJdOjpTaG93V2luZG93KFtXaW5kb3dNYW5hZ2VyXTo6R2V0Q29uc29sZVdpbmRvdygpLCAwKQ0KDQokRXhjbHVzaW9ucyA9IEAoDQogICAgIklkbGUiLCAiU3lzdGVtIiwgIlJlZ2lzdHJ5IiwgInNtc3MiLCAiY3Nyc3MiLCAid2luaW5pdCIsICJzZXJ2aWNlcyIsIA0KICAgICJsc2FzcyIsICJ3aW5sb2dvbiIsICJmb250ZHJ2aG9zdCIsICJkd20iLCAiTWVtb3J5IENvbXByZXNzaW9uIiwgIk1zTXBFbmciDQopDQoNCndoaWxlICgkdHJ1ZSkgew0KICAgIHRyeSB7DQogICAgICAgICMgRmluZCBIaWdoIE1lbW9yeSBQcm9jZXNzZXMgKD4xMDBNQikNCiAgICAgICAgJHRhcmdldHMgPSBHZXQtUHJvY2VzcyAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZSB8IFdoZXJlLU9iamVjdCB7IA0KICAgICAgICAgICAgJF8uV29ya2luZ1NldCAtZ3QgMTAwTUIgLWFuZCANCiAgICAgICAgICAgICRfLlByb2Nlc3NOYW1lIC1ub3RpbiAkRXhjbHVzaW9ucyANCiAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgaWYgKCR0YXJnZXRzKSB7DQogICAgICAgICAgICBmb3JlYWNoICgkcHJvYyBpbiAkdGFyZ2V0cykgew0KICAgICAgICAgICAgICAgIFtNZW1vcnlUcmltbWVyXTo6VHJpbVByb2Nlc3MoJHByb2MuSWQpDQogICAgICAgICAgICB9DQogICAgICAgIH0NCiAgICB9DQogICAgY2F0Y2ggeyB9DQogICAgDQogICAgU3RhcnQtU2xlZXAgLVNlY29uZHMgNQ0KfQ0KJ0ANCg0KIyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0NCiMgICAgICAgICAgICAgTUFJTiBMT0dJQw0KIyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0NCg0KZnVuY3Rpb24gU2hvdy1NZW51IHsNCiAgICBDbGVhci1Ib3N0DQogICAgV3JpdGUtSG9zdCAiPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09IiAtRm9yZWdyb3VuZENvbG9yIEN5YW4NCiAgICBXcml0ZS1Ib3N0ICIgICBHTE9CQUwgUkFNIE9QVElNSVpFUiAoU01BUlQgTU9ERSkiIC1Gb3JlZ3JvdW5kQ29sb3IgQ3lhbg0KICAgIFdyaXRlLUhvc3QgIj09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSIgLUZvcmVncm91bmRDb2xvciBDeWFuDQogICAgV3JpdGUtSG9zdCAiMS4gVGVzdCBHbG9iYWwgT3B0aW1pemF0aW9uIChWaXNpYmxlKSIgLUZvcmVncm91bmRDb2xvciBZZWxsb3cNCiAgICBXcml0ZS1Ib3N0ICIyLiBJbnN0YWxsIEdsb2JhbCBQZXJtYW5lbnQgKFNpbGVudCkiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4NCiAgICBXcml0ZS1Ib3N0ICIzLiBSZW1vdmUgR2xvYmFsIE9wdGltaXphdGlvbiIgLUZvcmVncm91bmRDb2xvciBSZWQNCiAgICBXcml0ZS1Ib3N0ICJRLiBRdWl0IiAtRm9yZWdyb3VuZENvbG9yIEdyYXkNCiAgICBXcml0ZS1Ib3N0ICI9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0iIC1Gb3JlZ3JvdW5kQ29sb3IgQ3lhbg0KfQ0KDQp3aGlsZSAoJHRydWUpIHsNCiAgICBTaG93LU1lbnUNCiAgICAkY2hvaWNlID0gUmVhZC1Ib3N0ICJTZWxlY3QgYW4gb3B0aW9uIg0KICAgIA0KICAgIHN3aXRjaCAoJGNob2ljZSkgew0KICAgICAgICAiMSIgew0KICAgICAgICAgICAgIyAtLS0gMS4gVEVTVCBNT0RFIC0tLQ0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiYG5MYXVuY2hpbmcgR2xvYmFsIFRlc3QgU2NyaXB0Li4uIiAtRm9yZWdyb3VuZENvbG9yIFllbGxvdw0KICAgICAgICAgICAgJHRlbXBTY3JpcHQgPSBKb2luLVBhdGggJGVudjpURU1QICJHbG9iYWwtT3B0aW1pemUtVGVzdC5wczEiDQogICAgICAgICAgICAkU2NyaXB0X1Rlc3RfR2xvYmFsIHwgT3V0LUZpbGUgLUZpbGVQYXRoICR0ZW1wU2NyaXB0IC1FbmNvZGluZyBVVEY4IC1Gb3JjZQ0KICAgICAgICAgICAgU3RhcnQtUHJvY2VzcyBwb3dlcnNoZWxsLmV4ZSAtQXJndW1lbnRMaXN0ICItTm9FeGl0IC1FeGVjdXRpb25Qb2xpY3kgQnlwYXNzIC1GaWxlIGAiJHRlbXBTY3JpcHRgIiINCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIkdsb2JhbCBPcHRpbWl6ZXIgcnVubmluZyBpbiBuZXcgd2luZG93LiIgLUZvcmVncm91bmRDb2xvciBHcmVlbg0KICAgICAgICAgICAgUGF1c2UNCiAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgIjIiIHsNCiAgICAgICAgICAgICMgLS0tIDIuIFBFUk1BTkVOVCBJTlNUQUxMQVRJT04gLS0tDQogICAgICAgICAgICBXcml0ZS1Ib3N0ICJgbkdlbmVyYXRpbmcgR2xvYmFsIE9wdGltaXphdGlvbiBTY3JpcHRzLi4uIiAtRm9yZWdyb3VuZENvbG9yIEN5YW4NCiAgICAgICAgICAgIA0KICAgICAgICAgICAgJEluc3RhbGxEaXIgPSAiQzpcR2xvYmFsUmFtT3B0aW1pemF0aW9uIg0KICAgICAgICAgICAgaWYgKC1ub3QgKFRlc3QtUGF0aCAkSW5zdGFsbERpcikpIHsgTmV3LUl0ZW0gLVBhdGggJEluc3RhbGxEaXIgLUl0ZW1UeXBlIERpcmVjdG9yeSAtRm9yY2UgfCBPdXQtTnVsbCB9DQogICAgICAgICAgICANCiAgICAgICAgICAgICRHbG9iYWxTY3JpcHQgPSAiJEluc3RhbGxEaXJcR2xvYmFsLU9wdGltaXplci5wczEiDQogICAgICAgICAgICAkU2NyaXB0X1Blcm1fR2xvYmFsIHwgT3V0LUZpbGUgLUZpbGVQYXRoICRHbG9iYWxTY3JpcHQgLUVuY29kaW5nIFVURjggLUZvcmNlDQogICAgICAgICAgICANCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIkZpbGUgZ2VuZXJhdGVkIGluICRJbnN0YWxsRGlyIiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuDQogICAgICAgICAgICANCiAgICAgICAgICAgICMgUmVnaXN0ZXIgVGFzaw0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiUmVnaXN0ZXJpbmcgU2NoZWR1bGVkIFRhc2suLi4iIC1Gb3JlZ3JvdW5kQ29sb3IgQ3lhbg0KICAgICAgICAgICAgJFRhc2tOYW1lID0gIkdsb2JhbFJhbU9wdGltaXplciINCiAgICAgICAgICAgIA0KICAgICAgICAgICAgVW5yZWdpc3Rlci1TY2hlZHVsZWRUYXNrIC1UYXNrTmFtZSAkVGFza05hbWUgLUNvbmZpcm06JGZhbHNlIC1FcnJvckFjdGlvbiBTaWxlbnRseUNvbnRpbnVlDQogICAgICAgICAgICAgICAgDQogICAgICAgICAgICAkQWN0aW9uID0gTmV3LVNjaGVkdWxlZFRhc2tBY3Rpb24gLUV4ZWN1dGUgInBvd2Vyc2hlbGwuZXhlIiAtQXJndW1lbnQgIi1XaW5kb3dTdHlsZSBIaWRkZW4gLUV4ZWN1dGlvblBvbGljeSBCeXBhc3MgLUZpbGUgYCIkR2xvYmFsU2NyaXB0YCIiDQogICAgICAgICAgICAkVHJpZ2dlciA9IE5ldy1TY2hlZHVsZWRUYXNrVHJpZ2dlciAtQXRMb2dPbg0KICAgICAgICAgICAgDQogICAgICAgICAgICAjIEZpeCBmb3IgIlBhcmFtZXRlciBpcyBpbmNvcnJlY3QiOiBVc2UgZnVsbHkgcXVhbGlmaWVkIHVzZXIgbmFtZSAoRE9NQUlOXFVzZXIpDQogICAgICAgICAgICAkVXNlciA9IFtTeXN0ZW0uU2VjdXJpdHkuUHJpbmNpcGFsLldpbmRvd3NJZGVudGl0eV06OkdldEN1cnJlbnQoKS5OYW1lDQogICAgICAgICAgICAkUHJpbmNpcGFsID0gTmV3LVNjaGVkdWxlZFRhc2tQcmluY2lwYWwgLVVzZXJJZCAkVXNlciAtTG9nb25UeXBlIEludGVyYWN0aXZlIC1SdW5MZXZlbCBIaWdoZXN0DQogICAgICAgICAgICANCiAgICAgICAgICAgICRTZXR0aW5ncyA9IE5ldy1TY2hlZHVsZWRUYXNrU2V0dGluZ3NTZXQgLUFsbG93U3RhcnRJZk9uQmF0dGVyaWVzIC1Eb250U3RvcElmR29pbmdPbkJhdHRlcmllcyAtRXhlY3V0aW9uVGltZUxpbWl0IChbVGltZVNwYW5dOjpaZXJvKQ0KICAgICAgICAgICAgICAgIA0KICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICBSZWdpc3Rlci1TY2hlZHVsZWRUYXNrIC1UYXNrTmFtZSAkVGFza05hbWUgLUFjdGlvbiAkQWN0aW9uIC1UcmlnZ2VyICRUcmlnZ2VyIC1QcmluY2lwYWwgJFByaW5jaXBhbCAtU2V0dGluZ3MgJFNldHRpbmdzIC1Gb3JjZSB8IE91dC1OdWxsDQogICAgICAgICAgICAgICAgDQogICAgICAgICAgICAgICAgIyBTVEFSVCBJTU1FRElBVEVMWQ0KICAgICAgICAgICAgICAgIFN0YXJ0LVNjaGVkdWxlZFRhc2sgLVRhc2tOYW1lICRUYXNrTmFtZQ0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlNVQ0NFU1M6IEdsb2JhbCBTbWFydCBPcHRpbWl6YXRpb24gSW5zdGFsbGVkICYgU3RhcnRlZCEiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJFcnJvciByZWdpc3RlcmluZyB0YXNrOiAkXyIgLUZvcmVncm91bmRDb2xvciBSZWQNCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIFBhdXNlDQogICAgICAgIH0NCiAgICAgICAgDQogICAgICAgICIzIiB7DQogICAgICAgICAgICAjIC0tLSAzLiBSRU1PVkUgQUxMIC0tLQ0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiYG5SZW1vdmluZyBHbG9iYWwgT3B0aW1pemF0aW9uLi4uIiAtRm9yZWdyb3VuZENvbG9yIFJlZA0KICAgICAgICAgICAgDQogICAgICAgICAgICAjIFN0b3AgUHJvY2Vzc2VzIChHbG9iYWwgKyBMZWdhY3kgU2NyaXB0cykNCiAgICAgICAgICAgIEdldC1XbWlPYmplY3QgV2luMzJfUHJvY2VzcyB8IFdoZXJlLU9iamVjdCB7ICRfLkNvbW1hbmRMaW5lIC1tYXRjaCAiR2xvYmFsLU9wdGltaXplcnxHbG9iYWwtT3B0aW1pemUtVGVzdHxPcHRpbWl6ZS1BbGx8T3B0aW1pemUtQ2hyb21lfE9wdGltaXplLVZTQ29kZXxPcHRpbWl6ZS1WaXN1YWxTdHVkaW8iIH0gfCBGb3JFYWNoLU9iamVjdCB7IA0KICAgICAgICAgICAgICAgIFN0b3AtUHJvY2VzcyAtSWQgJF8uUHJvY2Vzc0lkIC1Gb3JjZSAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZSANCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIlN0b3BwZWQgcnVubmluZyBzY3JpcHRzIChHbG9iYWwgJiBMZWdhY3kpLiIgLUZvcmVncm91bmRDb2xvciBHcmF5DQogICAgICAgICAgICANCiAgICAgICAgICAgICMgVW5yZWdpc3RlciBUYXNrcw0KICAgICAgICAgICAgVW5yZWdpc3Rlci1TY2hlZHVsZWRUYXNrIC1UYXNrTmFtZSAiR2xvYmFsUmFtT3B0aW1pemVyIiAtQ29uZmlybTokZmFsc2UgLUVycm9yQWN0aW9uIFNpbGVudGx5Q29udGludWUNCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIlVucmVnaXN0ZXJlZCB0YXNrcy4iIC1Gb3JlZ3JvdW5kQ29sb3IgR3JheQ0KDQogICAgICAgICAgICAjIERlbGV0ZSBGaWxlcw0KICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAiQzpcR2xvYmFsUmFtT3B0aW1pemF0aW9uIikgew0KICAgICAgICAgICAgICAgIFJlbW92ZS1JdGVtICJDOlxHbG9iYWxSYW1PcHRpbWl6YXRpb24iIC1SZWN1cnNlIC1Gb3JjZSAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZQ0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkRlbGV0ZWQgQzpcR2xvYmFsUmFtT3B0aW1pemF0aW9uIiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuDQogICAgICAgICAgICB9DQogICAgICAgICAgICANCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIkNsZWFudXAgQ29tcGxldGUuIiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuDQogICAgICAgICAgICBQYXVzZQ0KICAgICAgICB9DQogICAgICAgIA0KICAgICAgICAiUSIgeyBFeGl0IH0NCiAgICAgICAgInEiIHsgRXhpdCB9DQogICAgfQ0KfQ0K"
    
    # Check if script exists or matches (optional cache check could go here)
    $tempScript = Join-Path $env:TEMP "Global-Ram-Optimizer-Embedded.ps1"
    
    # Decode and Write
    try {
        $bytes = [System.Convert]::FromBase64String($b64)
        [System.IO.File]::WriteAllBytes($tempScript, $bytes)
        
        Write-Host "✅ Extracted to: $tempScript" -ForegroundColor DarkGray
        Write-Host "🚀 Launching..." -ForegroundColor Green
        
        # Run it
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`""
    }
    catch {
        Write-Host "❌ Error extracting/running script: $($_.Exception.Message)" -ForegroundColor Red
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
    $choice = Read-Host "`nEnter your choice [0-6]"
    switch ($choice) {
        '1' { Install-NormalSoftware }
        '2' { Install-MSOffice }
        '3' { Invoke-Activation }
        '4' { Update-AllSoftware }
        '5' { Invoke-AdvancedToolkit }
        '6' {
            Launch-GlobalOptimizer
        }
        '0' {
            Write-Host "`n👋 Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            break
        }
        default {
            Show-Menu -StatusMessage "⚠️ Invalid selection! Please choose between 0-6." -StatusColor "Red"
        }
    }
} while ($true)

