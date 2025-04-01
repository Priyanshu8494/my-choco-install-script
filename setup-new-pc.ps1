<#
.SYNOPSIS
  Priyanshu Suryavanshi PC Setup Toolkit
.DESCRIPTION
  Automated PC setup with software installation and system activation
.NOTES
  - Work in progress.
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
    
    # Display the .NOTES section here
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "                Work in Progress Note                      " -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " - Work in progress." -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    
    if ($StatusMessage) {
        Write-Host "[STATUS] $StatusMessage" -ForegroundColor $StatusColor
    }

    Write-Host "Main Menu Options:" -ForegroundColor Green
    Write-Host "1. Install Essential Software" -ForegroundColor White
    Write-Host "2. Install MS Office Suite" -ForegroundColor White
    Write-Host "3. System Activation Toolkit (Windows & Office)" -ForegroundColor White
    Write-Host "4. Update All Installed Software (Using Winget)" -ForegroundColor White
    Write-Host "0. Exit" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Cyan
}

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
    
    Write-Host "Select software to install (Enter numbers separated by commas):" -ForegroundColor Yellow
    for ($i = 0; $i -lt $softwareList.Count; $i++) {
        Write-Host "  $($i + 1). $($softwareList[$i].Name)" -ForegroundColor White
    }
    
    $selection = Read-Host "Enter selection (e.g., 1,3,5)"
    $selectedIndices = $selection -split "," | ForEach-Object {$_ -as [int]}
    
    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $softwareList.Count) {
            $app = $softwareList[$index - 1]
            Write-Host "  Installing $($app.Name)..." -ForegroundColor Gray
            winget install --id=$($app.ID) --silent --accept-source-agreements --accept-package-agreements
        } else {
            Write-Host "  Invalid selection: $index" -ForegroundColor Red
        }
    }
    
    Write-Host "`n✅ Selected software installed successfully!" -ForegroundColor Green
    Read-Host "`nPress Enter to return to the menu..."
}

function Install-MSOffice {
    <#
    .SYNOPSIS
    Download Office 2019, 2021, 2024, and 365

    .PARAMETER Branch
    Choose Office branch: 2019, 2021, 2024, and 365

    .PARAMETER Channel
    Choose Office channel: 2019, 2021, 2024, and 365

    .PARAMETER Components
    Choose Office components: Access, OneDrive, Outlook, Word, Excel, PowerPoint, Teams, OneNote, Publisher, Project 2019/2021/2024

    .EXAMPLE Download Office 2019 with the Word, Excel, PowerPoint components
    Download.ps1 -Branch ProPlus2019Retail -Channel Current -Components Word, Excel, PowerPoint

    .EXAMPLE Download Office 2021 with the Excel, Word components
    Download.ps1 -Branch ProPlus2021Volume -Channel PerpetualVL2021 -Components Excel, Word

    .EXAMPLE Download Office 2024 with the Excel, Word components
    Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Excel, Word

    .EXAMPLE Download Office 365 with the Excel, Word, PowerPoint components
    Download.ps1 -Branch O365ProPlusRetail -Channel Current -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word

    .LINK
    https://config.office.com/deploymentsettings

    .LINK
    https://docs.microsoft.com/en-us/deployoffice/vlactivation/gvlks

    .NOTES
    Run as non-admin
    #>

    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("ProPlus2019Retail", "ProPlus2021Volume", "ProPlus2024Volume", "O365ProPlusRetail")]
        [string]$Branch,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Current", "PerpetualVL2021", "PerpetualVL2024", "SemiAnnual")]
        [string]$Channel,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Access", "OneDrive", "Outlook", "Word", "Excel", "OneNote", "Publisher", "PowerPoint", "Teams", "ProjectPro2019Volume", "ProjectPro2021Volume", "ProjectPro2024Volume")]
        [string[]]$Components
    )

    if (-not (Test-Path -Path "$PSScriptRoot\Default.xml"))
    {
        Write-Warning -Message "Default.xml doesn't exist"
        exit
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if ($Host.Version.Major -eq 5)
    {
        $Script:ProgressPreference = "SilentlyContinue"
    }

    [xml]$Config = Get-Content -Path "$PSScriptRoot\Default.xml" -Encoding Default -Force

    switch ($Branch)
    {
        ProPlus2019Retail { ($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "ProPlus2019Retail" }
        ProPlus2021Volume { ($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "ProPlus2021Volume" }
        ProPlus2024Volume { ($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "ProPlus2024Volume" }
        O365ProPlusRetail { ($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "O365ProPlusRetail" }
    }

    switch ($Channel)
    {
        Current { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "Current" }
        PerpetualVL2021 { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "PerpetualVL2021" }
        PerpetualVL2024 { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "PerpetualVL2024" }
        SemiAnnual { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "SemiAnnual" }
    }

    foreach ($Component in $Components)
    {
        switch ($Component)
        {
            Access { $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Access']"); $Node.ParentNode.RemoveChild($Node) }
            Excel { $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Excel']"); $Node.ParentNode.RemoveChild($Node) }
            OneDrive { $OneDrive = Get-Package -Name "Microsoft OneDrive" -ProviderName Programs -Force -ErrorAction Ignore; if (-not $OneDrive) { Start-Process -FilePath $env:SystemRoot\SysWOW64\OneDriveSetup.exe } }
            Outlook { $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Outlook']"); $Node.ParentNode.RemoveChild($Node) }
            Word { $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Word']"); $Node.ParentNode.RemoveChild($Node) }
            PowerPoint { $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='PowerPoint']"); $Node.ParentNode.RemoveChild($Node) }
            Teams { $Parameters = @{ Uri = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"; OutFile = "$PSScriptRoot\MSTeams-x64.msix" }; Invoke-RestMethod @Parameters }
        }
    }

    $Config.Save("$PSScriptRoot\Config.xml")
    Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/download `"$PSScriptRoot\Config.xml`"" -Wait
    Write-Host "✅ Office download completed!" -ForegroundColor Green
    Read-Host "`nPress Enter to return to the menu..."
}

function Invoke-Activation {
    try {
        Write-Host "Activating Windows & Office..." -ForegroundColor Yellow
        irm https://get.activated.win | iex
        Write-Host "`n✅ Activation completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Activation failed: $_" -ForegroundColor Red
    }
    Read-Host "`nPress Enter to return to the menu..."
}

function Update-AllSoftware {
    try {
        Write-Host "Updating all installed software using Winget..." -ForegroundColor Yellow
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
        Write-Host "`n✅ All software updated successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Update failed: $_" -ForegroundColor Red
    }
    Read-Host "`nPress Enter to return to the menu..."
}

# Main program flow
do {
    Show-Menu
    $choice = Read-Host "`nEnter your choice [0-4]"

    switch ($choice) {
        '1' {
            Install-NormalSoftware
        }
        '2' {
            Install-MSOffice
        }
        '3' {
            Invoke-Activation
        }
        '4' {
            Update-AllSoftware
        }
        '0' { 
            Write-Host "Thank you for using Priyanshu Suryavanshi PC Setup Toolkit!" -ForegroundColor Cyan
            exit 
        }
        default {
            Show-Menu -StatusMessage "Invalid selection! Please choose between 0-4" -StatusColor "Red"
        }
    }
} while ($true)
