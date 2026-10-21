<#
.SYNOPSIS
    Installs the PowerShell modules needed for the OAL deployment, and optionally registers
    the Entra app that PnP.PowerShell needs for interactive sign-in.

.DESCRIPTION
    Run in PowerShell 7.4+ (PnP.PowerShell 2.x/3.x requires it).
    - PnP.PowerShell                          : site, libraries, permissions, uploads
    - Microsoft.Online.SharePoint.PowerShell  : Add-SPOOrgAssetsLibrary (only available here)

    Since September 2024 PnP no longer ships a shared multi-tenant app; every tenant needs its own
    app registration. -RegisterPnPApp creates one (needs Global Admin or Application Admin consent).

.EXAMPLE
    .\00-Install-Prerequisites.ps1
.EXAMPLE
    .\00-Install-Prerequisites.ps1 -RegisterPnPApp -Tenant contoso.onmicrosoft.com
#>
[CmdletBinding()]
param(
    [switch]$RegisterPnPApp,
    [string]$Tenant,                       # contoso.onmicrosoft.com
    [string]$AppName = 'PnP PowerShell - OAL deployment'
)
$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion -lt [version]'7.4') {
    throw "PowerShell 7.4 or later is required (this is $($PSVersionTable.PSVersion)). Install from https://aka.ms/powershell"
}

foreach ($m in 'PnP.PowerShell', 'Microsoft.Online.SharePoint.PowerShell') {
    $installed = Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1
    if ($installed) {
        Write-Host "$m $($installed.Version) already installed. Updating if a newer version exists..."
        Update-Module -Name $m -ErrorAction SilentlyContinue
    } else {
        Write-Host "Installing $m..."
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
}

if ($RegisterPnPApp) {
    if (-not $Tenant) { throw 'Pass -Tenant contoso.onmicrosoft.com with -RegisterPnPApp.' }
    Import-Module PnP.PowerShell
    # Creates an app with delegated SharePoint + Graph permissions and opens a consent prompt.
    $app = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $AppName -Tenant $Tenant -Interactive
    Write-Host "`nApp registered. Put this in OalConfig.psd1 -> PnPClientId:" -ForegroundColor Green
    $app | Format-List
}

Write-Host "`nDone. Next: fill in OalConfig.psd1, then run 01-New-AssetSite.ps1" -ForegroundColor Green
