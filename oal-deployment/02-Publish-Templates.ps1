<#
.SYNOPSIS
    Uploads templates (and optionally images) from a local folder tree to the asset libraries,
    recreating the folder structure (e.g. Company A\Consulting\Report.dotx).

.DESCRIPTION
    - Only .dotx / .potx / .xltx go to the template library (the OAL ignores other formats;
      macro-enabled .dotm/.potm/.xltm are not supported).
    - Existing files are overwritten (a new version is kept by library versioning).
    - With content approval on, uploaded files are approved immediately (-Approve).
    - Supports -WhatIf to preview.

.EXAMPLE
    .\02-Publish-Templates.ps1 -WhatIf
.EXAMPLE
    .\02-Publish-Templates.ps1 -IncludeImages
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'OalConfig.psd1'),
    [switch]$IncludeImages
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OalCommon.ps1')
$cfg = Get-OalConfig -Path $ConfigPath

$templateExt = '.dotx', '.potx', '.xltx'
$imageExt    = '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg'
$notSupported = '.docx', '.pptx', '.xlsx', '.dotm', '.potm', '.xltm', '.doc', '.ppt', '.xls', '.dot', '.pot', '.xlt'

function Publish-OalFolder {
    param([string]$LocalRoot, [string]$LibraryUrl, [string[]]$Extensions, [bool]$Approve)

    if (-not (Test-Path $LocalRoot)) { throw "Local folder not found: $LocalRoot" }
    $root = (Resolve-Path $LocalRoot).Path.TrimEnd('\', '/')
    $stats = @{ Uploaded = 0; Skipped = 0; Folders = 0 }

    # Folders first, so empty company/service-line folders exist too.
    Get-ChildItem -Path $root -Directory -Recurse | Sort-Object FullName | ForEach-Object {
        $rel = $_.FullName.Substring($root.Length).TrimStart('\', '/') -replace '\\', '/'
        $target = "$LibraryUrl/$rel"
        if ($PSCmdlet.ShouldProcess($target, 'Ensure folder')) {
            Resolve-PnPFolder -SiteRelativePath $target | Out-Null
        }
        $stats.Folders++
    }

    Get-ChildItem -Path $root -File -Recurse | Sort-Object FullName | ForEach-Object {
        $ext = $_.Extension.ToLowerInvariant()
        $relDir = $_.DirectoryName.Substring($root.Length).TrimStart('\', '/') -replace '\\', '/'
        $folder = if ($relDir) { "$LibraryUrl/$relDir" } else { $LibraryUrl }

        if ($Extensions -notcontains $ext) {
            if ($notSupported -contains $ext) { Write-OalWarn "$($_.Name): $ext not supported here. Save it as $($Extensions -join ' / ')." }
            else { Write-OalSkip "$($_.Name) (not a $($Extensions -join '/') file)" }
            $stats.Skipped++
            return
        }
        if ($PSCmdlet.ShouldProcess("$folder/$($_.Name)", 'Upload')) {
            if ($Approve) { Add-PnPFile -Path $_.FullName -Folder $folder -Approve -ApproveComment 'Published by OAL deployment script' | Out-Null }
            else          { Add-PnPFile -Path $_.FullName -Folder $folder | Out-Null }
            Write-OalOk "$folder/$($_.Name)"
        }
        $stats.Uploaded++
    }
    return $stats
}

Connect-OalPnP -Url $cfg.SiteUrl -ClientId $cfg.PnPClientId
$approve = [bool]$cfg.EnableContentApproval

Write-OalStep "Templates: $($cfg.LocalTemplateRoot) -> /$($cfg.TemplateLibraryUrl)"
$t = Publish-OalFolder -LocalRoot $cfg.LocalTemplateRoot -LibraryUrl $cfg.TemplateLibraryUrl -Extensions $templateExt -Approve $approve
Write-Host "    $($t.Uploaded) uploaded, $($t.Skipped) skipped, $($t.Folders) folders"

if ($IncludeImages) {
    if (-not $cfg.ImageLibrary.Enabled) { throw 'ImageLibrary.Enabled is $false in the config.' }
    Write-OalStep "Images: $($cfg.ImageLibrary.LocalRoot) -> /$($cfg.ImageLibrary.Url)"
    $i = Publish-OalFolder -LocalRoot $cfg.ImageLibrary.LocalRoot -LibraryUrl $cfg.ImageLibrary.Url -Extensions $imageExt -Approve $approve
    Write-Host "    $($i.Uploaded) uploaded, $($i.Skipped) skipped, $($i.Folders) folders"
}

Write-Host "`nDone. Next (first time only): 03-Register-OrgAssetsLibrary.ps1" -ForegroundColor Green
