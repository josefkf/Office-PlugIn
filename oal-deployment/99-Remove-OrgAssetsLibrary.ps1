<#
.SYNOPSIS
    Rollback: unregisters the OAL libraries and removes their CDN origins. Optionally deletes the site.

.DESCRIPTION
    Removing a library from the OAL leaves its CDN origin in place, so this also removes the origin.
    The site and its files are kept unless -DeleteSite is passed (site goes to the recycle bin).

.EXAMPLE
    .\99-Remove-OrgAssetsLibrary.ps1 -WhatIf
.EXAMPLE
    .\99-Remove-OrgAssetsLibrary.ps1 -DeleteSite
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'OalConfig.psd1'),
    [switch]$DeleteSite
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OalCommon.ps1')
$cfg = Get-OalConfig -Path $ConfigPath

Connect-OalSpo -AdminUrl $cfg.AdminUrl

$urls = @($cfg.TemplateLibraryAbsUrl)
if ($cfg.ImageLibrary.Enabled) { $urls += $cfg.ImageLibraryAbsUrl }

$registered = @(Get-OalRegisteredLibraries)
# Origins can carry a " (configuration pending)" suffix; strip it before comparing.
$origins = @(Get-SPOTenantCdnOrigins -CdnType Private) | ForEach-Object { ("$_" -replace '\s*\(.*\)\s*$', '').Trim() }

foreach ($u in $urls) {
    $rel = Get-OalRelPath $u
    Write-OalStep $u
    if ($registered | Where-Object Url -eq $rel) {
        if ($PSCmdlet.ShouldProcess($u, 'Remove-SPOOrgAssetsLibrary')) {
            Remove-SPOOrgAssetsLibrary -LibraryUrl $u -Confirm:$false
            Write-OalOk 'Unregistered'
        }
    } else { Write-OalSkip 'Not registered' }

    $origin = $origins | Where-Object { (Get-OalRelPath $_) -eq $rel } | Select-Object -First 1
    if ($origin) {
        if ($PSCmdlet.ShouldProcess($origin, 'Remove-SPOTenantCdnOrigin')) {
            Remove-SPOTenantCdnOrigin -CdnType Private -OriginUrl $origin -Confirm:$false
            Write-OalOk "CDN origin $origin removed"
        }
    } else { Write-OalSkip 'No CDN origin' }
}

if ($DeleteSite) {
    Write-OalStep "Delete site $($cfg.SiteUrl)"
    if ($PSCmdlet.ShouldProcess($cfg.SiteUrl, 'Remove-SPOSite (to recycle bin)')) {
        Remove-SPOSite -Identity $cfg.SiteUrl -Confirm:$false
        Write-OalOk 'Site moved to the recycle bin (restorable for 93 days)'
    }
}

Write-Host "`nDone. Changes can take up to 24h to disappear from Office apps." -ForegroundColor Green
