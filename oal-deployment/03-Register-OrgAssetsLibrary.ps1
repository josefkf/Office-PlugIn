<#
.SYNOPSIS
    Registers the template library (and optional image library) as Organization Assets Libraries,
    then verifies the setup.

.DESCRIPTION
    Uses Add-SPOOrgAssetsLibrary, which also enables the tenant CDN for the library (no separate
    CDN step needed). Idempotent: already-registered libraries are skipped.
    Templates can take up to 24 hours to appear in Office apps.

    Requires: SharePoint Administrator.

.EXAMPLE
    .\03-Register-OrgAssetsLibrary.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param([string]$ConfigPath = (Join-Path $PSScriptRoot 'OalConfig.psd1'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OalCommon.ps1')
$cfg = Get-OalConfig -Path $ConfigPath

Write-OalStep "Connecting to $($cfg.AdminUrl)"
Connect-OalSpo -AdminUrl $cfg.AdminUrl

$thumbUrl = if ($cfg.ThumbnailPath) { "$($cfg.SiteUrl)/AssetThumbnails/$(Split-Path $cfg.ThumbnailPath -Leaf)" } else { $null }

$targets = @(@{ Url = $cfg.TemplateLibraryAbsUrl; Type = 'OfficeTemplateLibrary' })
if ($cfg.ImageLibrary.Enabled) { $targets += @{ Url = $cfg.ImageLibraryAbsUrl; Type = 'ImageDocumentLibrary' } }

$registered = @(Get-OalRegisteredLibraries)
foreach ($t in $targets) {
    Write-OalStep "$($t.Type): $($t.Url)"
    if ($registered | Where-Object Url -eq (Get-OalRelPath $t.Url)) {
        Write-OalSkip 'Already registered'
        continue
    }
    $params = @{ LibraryUrl = $t.Url; OrgAssetType = $t.Type; Confirm = $false }
    if ($thumbUrl) { $params.ThumbnailUrl = $thumbUrl }
    if ($null -ne $cfg.CopilotSearchable) { $params.CopilotSearchable = [bool]$cfg.CopilotSearchable }
    if ($PSCmdlet.ShouldProcess($t.Url, "Add-SPOOrgAssetsLibrary ($($t.Type))")) {
        Add-SPOOrgAssetsLibrary @params
        Write-OalOk 'Registered'
    }
}

# --- Verify ----------------------------------------------------------------
Write-OalStep 'Verification'
$registered = @(Get-OalRegisteredLibraries)
foreach ($t in $targets) {
    $hit = $registered | Where-Object Url -eq (Get-OalRelPath $t.Url)
    if ($hit) { Write-OalOk "$($t.Url) registered as $($hit.Type)" }
    elseif (-not $WhatIfPreference) { Write-OalWarn "$($t.Url) NOT found in Get-SPOOrgAssetsLibrary" }
}
Write-Host '    Private CDN status:'
Get-SPOTenantCdnEnabled -CdnType Private | Out-String | ForEach-Object { Write-Host ($_.Trim() -replace '(?m)^', '      ') }

Write-Host @"

Done. Notes:
 - Allow up to 24h for templates to show in Word/Excel/PowerPoint (File > New > your organization tab).
 - Desktop apps need Microsoft 365 Apps 2002+. PowerPoint on the web needs an E3/E5 licence.
 - Optional: Intune policy to hide Office-provided templates, so users only see yours.
"@ -ForegroundColor Green
