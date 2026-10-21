<#
.SYNOPSIS
    Creates the asset site and libraries for the Organization Assets Library (OAL), with permissions.

.DESCRIPTION
    Idempotent: safe to re-run. Existing site/libraries are reused and settings re-applied.
    - Communication site (SiteUrl)
    - Template library (+ optional image library), versioning, optional content approval
    - "Everyone except external users" added to site Visitors (required by the OAL)
    - Editors (UPNs / Entra security groups) added to site Members
    - Optional thumbnail uploaded to an 'AssetThumbnails' library (must be on the same site)

    Requires: SharePoint Administrator. Run in PowerShell 7.4+.

.EXAMPLE
    .\01-New-AssetSite.ps1
#>
[CmdletBinding()]
param([string]$ConfigPath = (Join-Path $PSScriptRoot 'OalConfig.psd1'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OalCommon.ps1')
$cfg = Get-OalConfig -Path $ConfigPath

# --- 1. Site ---------------------------------------------------------------
Write-OalStep "Asset site $($cfg.SiteUrl)"
Connect-OalPnP -Url $cfg.AdminUrl -ClientId $cfg.PnPClientId
$site = Get-PnPTenantSite -Identity $cfg.SiteUrl -ErrorAction SilentlyContinue
if ($site) {
    Write-OalSkip 'Site already exists'
} else {
    New-PnPSite -Type CommunicationSite -Title $cfg.SiteTitle -Url $cfg.SiteUrl -Owner $cfg.SiteOwner -Wait | Out-Null
    Write-OalOk 'Site created'
}

Connect-OalPnP -Url $cfg.SiteUrl -ClientId $cfg.PnPClientId

# --- 2. Libraries ----------------------------------------------------------
function Initialize-OalLibrary {
    param([string]$Title, [string]$Url, [bool]$Approval)
    $list = Get-PnPList -Identity $Url -ErrorAction SilentlyContinue
    if (-not $list) { $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue }
    if ($list) {
        Write-OalSkip "Library '$Title' exists"
    } else {
        $list = New-PnPList -Title $Title -Url $Url -Template DocumentLibrary -OnQuickLaunch
        Write-OalOk "Library '$Title' created (/$Url)"
    }
    Set-PnPList -Identity $list -EnableVersioning $true -MajorVersions $cfg.MajorVersionLimit -EnableModeration $Approval | Out-Null
    Write-OalOk "Versioning on (max $($cfg.MajorVersionLimit)), content approval $(if ($Approval) {'on'} else {'off'})"
}

Write-OalStep 'Libraries'
Initialize-OalLibrary -Title $cfg.TemplateLibraryTitle -Url $cfg.TemplateLibraryUrl -Approval ([bool]$cfg.EnableContentApproval)
if ($cfg.ImageLibrary.Enabled) {
    Initialize-OalLibrary -Title $cfg.ImageLibrary.Title -Url $cfg.ImageLibrary.Url -Approval ([bool]$cfg.EnableContentApproval)
}

# --- 3. Permissions --------------------------------------------------------
Write-OalStep 'Permissions'
$tenantId = Get-PnPTenantId
$everyone = "c:0-.f|rolemanager|spo-grid-all-users/$tenantId"   # Everyone except external users
$visitors = Get-PnPGroup -AssociatedVisitorGroup
$members  = Get-PnPGroup -AssociatedMemberGroup

$visitorLogins = @(Get-PnPGroupMember -Group $visitors | ForEach-Object LoginName)
if ($visitorLogins -contains $everyone) {
    Write-OalSkip "'Everyone except external users' already in $($visitors.Title)"
} else {
    Add-PnPGroupMember -Group $visitors -LoginName $everyone
    Write-OalOk "'Everyone except external users' added to $($visitors.Title)"
}

$memberLogins = @(Get-PnPGroupMember -Group $members | ForEach-Object LoginName)
$editors = @($cfg.EditorUpns | Where-Object { $_ }) + @($cfg.EditorGroupObjectIds | Where-Object { $_ } | ForEach-Object { "c:0t.c|tenant|$_" })
foreach ($e in $editors) {
    if ($memberLogins | Where-Object { $_ -like "*$e" }) { Write-OalSkip "$e already in $($members.Title)"; continue }
    Add-PnPGroupMember -Group $members -LoginName $e
    Write-OalOk "$e added to $($members.Title)"
}
if (-not $editors) { Write-OalWarn 'No editors configured. Only site owners can upload templates.' }

# --- 4. Thumbnail ----------------------------------------------------------
if ($cfg.ThumbnailPath) {
    Write-OalStep 'Thumbnail'
    if (-not (Test-Path $cfg.ThumbnailPath)) { throw "ThumbnailPath not found: $($cfg.ThumbnailPath)" }
    if (-not (Get-PnPList -Identity 'AssetThumbnails' -ErrorAction SilentlyContinue)) {
        New-PnPList -Title 'AssetThumbnails' -Url 'AssetThumbnails' -Template DocumentLibrary | Out-Null
    }
    $f = Add-PnPFile -Path $cfg.ThumbnailPath -Folder 'AssetThumbnails'
    Write-OalOk "Uploaded $($f.Name)"
}

Write-Host "`nDone. Next: put templates in $($cfg.LocalTemplateRoot) and run 02-Publish-Templates.ps1" -ForegroundColor Green
