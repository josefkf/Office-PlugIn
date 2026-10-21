# Shared helpers, dot-sourced by the numbered scripts.

function Get-OalConfig {
    param([string]$Path = (Join-Path $PSScriptRoot 'OalConfig.psd1'))
    if (-not (Test-Path $Path)) { throw "Config not found: $Path" }
    $cfg = Import-PowerShellDataFile -Path $Path
    foreach ($k in 'TenantName', 'PnPClientId', 'SiteUrl', 'SiteTitle', 'SiteOwner', 'TemplateLibraryTitle', 'TemplateLibraryUrl') {
        if ([string]::IsNullOrWhiteSpace($cfg[$k])) { throw "Config value '$k' is empty in $Path" }
    }
    if ($cfg.PnPClientId -eq '00000000-0000-0000-0000-000000000000') {
        throw 'Set PnPClientId in OalConfig.psd1 (run 00-Install-Prerequisites.ps1 -RegisterPnPApp if you need one).'
    }
    $cfg.SiteUrl = $cfg.SiteUrl.TrimEnd('/')
    $cfg.AdminUrl = "https://$($cfg.TenantName)-admin.sharepoint.com"
    $cfg.TemplateLibraryAbsUrl = "$($cfg.SiteUrl)/$($cfg.TemplateLibraryUrl)"
    if ($cfg.ImageLibrary -and $cfg.ImageLibrary.Enabled) {
        $cfg.ImageLibraryAbsUrl = "$($cfg.SiteUrl)/$($cfg.ImageLibrary.Url)"
    }
    return $cfg
}

function Connect-OalPnP {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][string]$ClientId)
    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive -ErrorAction Stop
}

function Connect-OalSpo {
    param([Parameter(Mandatory)][string]$AdminUrl)
    # The SPO module is built for Windows PowerShell; in PowerShell 7 load it through the compatibility layer.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop
    } else {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
    }
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
}

# Server-relative path of a library URL, lower-cased, no trailing slash: /sites/x/templates
function Get-OalRelPath {
    param([Parameter(Mandatory)][string]$Url)
    $p = if ($Url -match '^https?://') { ([uri]$Url).AbsolutePath } else { $Url }
    return ('/' + [uri]::UnescapeDataString($p).Trim('/')).ToLowerInvariant()
}

# Registered OAL entries as plain objects: Url (server-relative, lower-case), Type, Raw
function Get-OalRegisteredLibraries {
    $raw = @(Get-SPOOrgAssetsLibrary -ErrorAction Stop) | Where-Object { $_ -and $_.PSObject.Properties['LibraryUrl'] }
    foreach ($lib in $raw) {
        $u = if ($lib.LibraryUrl.PSObject.Properties['DecodedUrl']) { $lib.LibraryUrl.DecodedUrl } else { "$($lib.LibraryUrl)" }
        [pscustomobject]@{ Url = (Get-OalRelPath $u); Type = "$($lib.OrgAssetType)"; Raw = $lib }
    }
}

function Write-OalStep { param([string]$Text) Write-Host "`n==> $Text" -ForegroundColor Cyan }
function Write-OalOk   { param([string]$Text) Write-Host "    OK   $Text" -ForegroundColor Green }
function Write-OalSkip { param([string]$Text) Write-Host "    SKIP $Text" -ForegroundColor DarkGray }
function Write-OalWarn { param([string]$Text) Write-Host "    WARN $Text" -ForegroundColor Yellow }
