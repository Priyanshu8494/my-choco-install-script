<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation

.NOTES
  - Work in progress.
#>

# Initialization and Requirements Check
$logFile = "$env:TEMP\PSSetupToolkit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$navigationStack = New-Object System.Collections.Stack
$wingetCache = @{}
$menuCache = $null

# Check if winget is available
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget is not available. Please install it first." -ForegroundColor Red
    exit 1
}

# Check for admin privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator" -ForegroundColor Red
    exit 1
}

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp][$Level] $Message"
    Add-Content -Path $logFile -Value $logEntry
    
    if ($Level -eq "ERROR") {
        Write-Host $Message -ForegroundColor Red
    }
}

# Retry function for unreliable operations
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$RetryDelay = 5
    )
    
    $attempt = 1
    while ($attempt -le $MaxRetries) {
        try {
            Write-Log "Attempt $attempt of $MaxRetries"
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $MaxRetries) {
                Write-Log "Final attempt failed: $_" -Level "ERROR"
                throw
            }
            Write-Log "Attempt $attempt failed. Retrying in $RetryDelay seconds... Error: $_" -Level "WARNING"
            Write-Host "Attempt $attempt failed. Retrying in $RetryDelay seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds $RetryDelay
            $attempt++
        }
    }
}

# Progress indicator
function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$Step,
        [int]$TotalSteps
    )
    $percentComplete = ($Step / $TotalSteps) * 100
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}

# Header and Menu functions
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "            Priyanshu Suryavanshi PC Setup Toolkit          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Menu {
    param (
        [string]$StatusMessage = "",
        [string]$StatusColor = "Yellow",
        [switch]$NoPush
    )
    
    if (-not $NoPush) {
        $navigationStack.Push((Get-PSCallStack)[0].ScriptName)
    }
    
    if (-not $menuCache) {
        $script:menuCache = {
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
        }.ToString()
    }
    
    Invoke-Expression $menuCache
}

function Go-Back {
    if ($navigationStack.Count -gt 1) {
        $null = $navigationStack.Pop()
        $previousMenu = $navigationStack.Peek()
        & $previousMenu -NoPush
    }
}

# Software Installation Functions
function Install-NormalSoftware {
    $softwareList = @(
        @{Name="Google Chrome"; ID="Google.Chrome"},
        @{Name="Mozilla Firefox"; ID="Mozilla.Firefox"},
        @{Name="WinRAR"; ID="RARLab.WinRAR"},
        @{Name="VLC Player"; ID="VideoLAN.VLC"},
        @{Name="PDF Reader"; ID="SumatraPDF.SumatraPDF"},
        @{Name="AnyDesk"; ID="AnyDeskSoftwareGmbH.AnyDesk"},
        @{Name="UltraViewer"; ID="UltraViewer.UltraViewer"}
    )
    
    # Cache winget results
    foreach ($app in $softwareList) {
        if (-not $wingetCache.ContainsKey($app.ID)) {
            $wingetCache[$app.ID] = (winget list --id $app.ID --accept-source-agreements) -like "*$($app.ID)*"
        }
    }
    
    Write-Host "Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        $status = if ($wingetCache[$softwareList[$i].ID]) { "[INSTALLED]" } else { "[NOT INSTALLED]" }
        Write-Host "  $($i + 1). $($softwareList[$i].Name) $status" -ForegroundColor White
    }
    
    $selection = Read-Host "`nEnter selection (e.g., 1,3,5 or 'all')"
    
    $selectedIndices = if ($selection -eq "all") {
        1..$softwareList.Count
    } else {
        $selection -split "," | ForEach-Object { $_ -as [int] }
    }
    
    $jobs = @()
    $toInstall = @()
    
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            if (-not $wingetCache[$app.ID]) {
                $toInstall += $app
            } else {
                Write-Host "  $($app.Name) is already installed. Skipping." -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    if ($toInstall.Count -gt 0) {
        # Parallel installation
        Write-Host "`nInstalling $($toInstall.Count) applications..." -ForegroundColor Yellow
        
        $installBlock = {
            param($id)
            winget install --id=$id --silent --accept-source-agreements --accept-package-agreements
        }
        
        $batchSize = [math]::Min(4, $toInstall.Count) # Limit to 4 parallel installs
        $batches = [math]::Ceiling($toInstall.Count / $batchSize)
        
        for ($batch = 0; $batch -lt $batches; $batch++) {
            $batchStart = $batch * $batchSize
            $batchEnd = [math]::Min(($batch + 1) * $batchSize - 1, $toInstall.Count - 1)
            
            $jobs = @()
            for ($i = $batchStart; $i -le $batchEnd; $i++) {
                $app = $toInstall[$i]
                Write-Host "  Starting installation of $($app.Name)..." -ForegroundColor Gray
                $jobs += Start-Job -ScriptBlock $installBlock -ArgumentList $app.ID -Name $app.Name
            }
            
            # Monitor jobs
            $completed = 0
            $totalJobs = $jobs.Count
            
            while ($jobs | Where-Object { $_.State -eq 'Running' }) {
                $newCompleted = ($jobs | Where-Object { $_.State -ne 'Running' }).Count
                if ($newCompleted -ne $completed) {
                    $completed = $newCompleted
                    Show-Progress -Activity "Installing Software" -Status "$completed of $totalJobs completed" -Step $completed -TotalSteps $totalJobs
                }
                Start-Sleep -Milliseconds 500
            }
            
            # Get results
            foreach ($job in $jobs) {
                $result = Receive-Job -Job $job
                if ($job.State -eq 'Failed') {
                    Write-Log "Installation failed for $($job.Name): $result" -Level "ERROR"
                } else {
                    $wingetCache[$job.Name] = $true
                    Write-Log "Successfully installed $($job.Name)"
                }
                Remove-Job -Job $job
            }
        }
    }
    
    Show-Menu -StatusMessage "✅ Selected software installation completed!" -StatusColor "Green"
}

function Install-MSOffice {
    try {
        Write-Host "`nDownloading MS Office setup..." -ForegroundColor Yellow
        $officeJob = Start-Job -ScriptBlock {
            try {
                Write-Log "Starting MS Office installation"
                Start-Process "msiexec.exe" -ArgumentList "/i", "C:\Path\To\OfficeSetup.msi", "/quiet", "/norestart" -Wait
                Write-Log "MS Office installation completed"
            } catch {
                Write-Log "Office installation error: $_" -Level "ERROR"
                throw
            }
        }
        
        while ($officeJob.State -eq 'Running') {
            Write-Host "." -NoNewline -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
        
        if ($officeJob.State -eq 'Failed') {
            throw (Receive-Job -Job $officeJob -ErrorAction SilentlyContinue)
        }
        
        Show-Menu -StatusMessage "✅ MS Office installed successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Office installation failed: $_" -StatusColor "Red"
    }
    finally {
        if ($officeJob) { Remove-Job -Job $officeJob }
    }
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        Invoke-WithRetry -ScriptBlock {
            irm https://get.activated.win | iex
        } -MaxRetries 3 -RetryDelay 10
        
        Show-Menu -StatusMessage "✅ Activation completed successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Activation failed: $_" -StatusColor "Red"
    }
}

function Update-AllSoftware {
    try {
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        $updateJob = Start-Job -ScriptBlock {
            winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        }
        
        while ($updateJob.State -eq 'Running') {
            Write-Host "." -NoNewline -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }
        
        if ($updateJob.State -eq 'Failed') {
            throw (Receive-Job -Job $updateJob -ErrorAction SilentlyContinue)
        }
        
        $output = Receive-Job -Job $updateJob
        Write-Log "Update output: $output"
        
        Show-Menu -StatusMessage "✅ All software updated successfully!" -StatusColor "Green"
    }
    catch {
        Show-Menu -StatusMessage "Update failed: $_" -StatusColor "Red"
    }
    finally {
        if ($updateJob) { Remove-Job -Job $updateJob }
    }
}

# Main program flow
try {
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
                exit 0 
            }
            default {
                Show-Menu -StatusMessage "Invalid selection! Please choose between 0-4" -StatusColor "Red"
            }
        }
    } while ($true)
}
catch {
    Write-Host "`nA critical error occurred: $_" -ForegroundColor Red
    Write-Host "Check the log file at $logFile for details" -ForegroundColor Yellow
    Write-Log "Critical error: $_" -Level "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Pause
    exit 1
}
