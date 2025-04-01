<#
    .SYNOPSIS
    Download Office 2024

    .PARAMETER Branch
    Choose Office branch: ProPlus2024Volume

    .PARAMETER Channel
    Choose Office channel: PerpetualVL2024 or Current

    .PARAMETER Components
    Choose Office components: Access, OneDrive, Outlook, Word, Excel, PowerPoint, Teams, OneNote, Publisher, ProjectPro2024Volume

    .EXAMPLE Download Office 2024 with the Word, Excel, PowerPoint components
    Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Word, Excel, PowerPoint

    .LINK
    https://config.office.com/deploymentsettings

    .LINK
    https://docs.microsoft.com/en-us/deployoffice/vlactivation/gvlks

    .NOTES
    Run as non-admin
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ProPlus2024Volume")]
    [string]$Branch,

    [Parameter(Mandatory = $true)]
    [ValidateSet("PerpetualVL2024", "Current")]
    [string]$Channel,

    [Parameter(Mandatory = $true)]
    [ValidateSet("Access", "OneDrive", "Outlook", "Word", "Excel", "OneNote", "Publisher", "PowerPoint", "Teams", "ProjectPro2024Volume")]
    [string[]]$Components
)

if (-not (Test-Path -Path "$PSScriptRoot\Default.xml")) {
    Write-Warning -Message "Default.xml doesn't exist"
    exit
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($Host.Version.Major -eq 5) {
    $Script:ProgressPreference = "SilentlyContinue"
}

[xml]$Config = Get-Content -Path "$PSScriptRoot\Default.xml" -Encoding Default -Force

switch ($Branch) {
    ProPlus2024Volume { ($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "ProPlus2024Volume" }
}

switch ($Channel) {
    Current { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "Current" }
    PerpetualVL2024 { ($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "PerpetualVL2024" }
}

foreach ($Component in $Components) {
    switch ($Component) {
        Access {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Access']")
            $Node.ParentNode.RemoveChild($Node)
        }
        Excel {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Excel']")
            $Node.ParentNode.RemoveChild($Node)
        }
        OneDrive {
            $OneDrive = Get-Package -Name "Microsoft OneDrive" -ProviderName Programs -Force -ErrorAction Ignore
            if (-not $OneDrive) {
                switch ((Get-CimInstance -ClassName Win32_OperatingSystem).Caption) {
                    {$_ -match 10} { if (Test-Path -Path $env:SystemRoot\SysWOW64\OneDriveSetup.exe) { Start-Process -FilePath $env:SystemRoot\SysWOW64\OneDriveSetup.exe } }
                    {$_ -match 11} { if (Test-Path -Path $env:SystemRoot\System32\OneDriveSetup.exe) { Start-Process -FilePath $env:SystemRoot\System32\OneDriveSetup.exe } }
                }
            }
        }
        Outlook {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Outlook']")
            $Node.ParentNode.RemoveChild($Node)
        }
        Word {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Word']")
            $Node.ParentNode.RemoveChild($Node)
        }
        PowerPoint {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='PowerPoint']")
            $Node.ParentNode.RemoveChild($Node)
        }
        OneNote {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='OneNote']")
            $Node.ParentNode.RemoveChild($Node)
        }
        Publisher {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Publisher']")
            $Node.ParentNode.RemoveChild($Node)
        }
        Teams {
            Write-Information -MessageData "" -InformationAction Continue
            $Parameters = @{
                Uri             = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
                OutFile         = "$env:USERPROFILE\Downloads\MSTeams-x64.msix"
                UseBasicParsing = $true
            }
            Invoke-RestMethod @Parameters
        }
        ProjectPro2024Volume {
            $ProjectNode = $Config.Configuration.Add.AppendChild($Config.CreateElement("Product"))
            $ProjectNode.SetAttribute("ID","ProjectPro2024Volume")
            $ProjectElement = $ProjectNode.AppendChild($Config.CreateElement("Language"))
            $ProjectElement.SetAttribute("ID","MatchOS")
        }
    }
}

$Config.Save("$PSScriptRoot\Config.xml")

if (((Get-WinHomeLocation).GeoId -eq "203") -or ((Get-WinHomeLocation).GeoId -eq "29")) {
    $Script:Region = (Get-WinHomeLocation).GeoId
    Set-WinHomeLocation -GeoId 241
    Write-Warning -Message "Region changed to Ukrainian"
    $Script:RegionChanged = $true
}

Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment -Recurse -Force -ErrorAction Ignore
Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs -Recurse -Force -ErrorAction Ignore
Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs -Recurse -Force -ErrorAction Ignore

if (-not (Test-Path -Path "$PSScriptRoot\setup.exe")) {
    $Parameters = @{
        Uri             = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
        OutFile         = "$PSScriptRoot\setup.exe"
        UseBasicParsing = $true
    }
    Invoke-WebRequest @Parameters
}

Write-Verbose -Message "Downloading... Please do not close any console windows." -Verbose
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/download `"$PSScriptRoot\Config.xml`"" -Wait

if ($Script:RegionChanged) {
    Set-WinHomeLocation -GeoId $Script:Region
    Write-Warning -Message "Region changed to original one"
}

Write-Verbose -Message "Office downloaded. Please run Install.ps1 file with administrator privileges." -Verbose
