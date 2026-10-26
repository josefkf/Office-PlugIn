<#
.SYNOPSIS
    Recursively converts .docx/.xlsx/.pptx files into .dotx/.xltx/.potx templates.

.PARAMETER Path
    Root folder to scan.
.PARAMETER OutputPath
    Optional. If set, templates are written there, mirroring the folder structure.
    If omitted, templates are created next to the originals.
.PARAMETER UseOffice
    Use Word/Excel/PowerPoint COM automation (SaveAs template). Requires Office installed.
    Default is a pure OpenXML approach (no Office needed): copy the file, rename the
    extension and patch the main part's content type in [Content_Types].xml.
.PARAMETER Overwrite
    Overwrite existing templates.

.EXAMPLE
    .\Convert-ToOfficeTemplates.ps1 -Path C:\Docs -OutputPath C:\Templates
.EXAMPLE
    .\Convert-ToOfficeTemplates.ps1 -Path C:\Docs -UseOffice -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$OutputPath,
    [switch]$UseOffice,
    [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Path = (Resolve-Path $Path).Path
if ($OutputPath) {
    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
    $OutputPath = (Resolve-Path $OutputPath).Path
}

# ext -> template ext, old main content type, new main content type, Office SaveAs format
$map = @{
    '.docx' = @{ Ext = '.dotx'; App = 'Word.Application'
                 Old = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
                 New = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
                 Fmt = 14 }  # wdFormatXMLTemplate
    '.xlsx' = @{ Ext = '.xltx'; App = 'Excel.Application'
                 Old = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
                 New = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
                 Fmt = 54 }  # xlOpenXMLTemplate
    '.pptx' = @{ Ext = '.potx'; App = 'PowerPoint.Application'
                 Old = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
                 New = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
                 Fmt = 26 }  # ppSaveAsOpenXMLTemplate
}

function Get-TargetPath([System.IO.FileInfo]$File, [string]$NewExt) {
    $name = [IO.Path]::ChangeExtension($File.Name, $NewExt)
    if ($OutputPath) {
        $rel = $File.DirectoryName.Substring($Path.Length).TrimStart('\','/')
        $dir = Join-Path $OutputPath $rel
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        return Join-Path $dir $name
    }
    return Join-Path $File.DirectoryName $name
}

function Convert-OpenXml([string]$Source, [string]$Target, $Info) {
    Copy-Item -LiteralPath $Source -Destination $Target -Force
    $zip = [IO.Compression.ZipFile]::Open($Target, 'Update')
    try {
        $entry = $zip.GetEntry('[Content_Types].xml')
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd(); $reader.Dispose()
        if ($xml -notmatch [regex]::Escape($Info.Old)) { throw "Main content type not found (macro-enabled or already a template?)" }
        $xml = $xml.Replace($Info.Old, $Info.New)
        $entry.Delete()
        $new = $zip.CreateEntry('[Content_Types].xml')
        $writer = New-Object IO.StreamWriter($new.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($xml); $writer.Dispose()
    }
    finally { $zip.Dispose() }
}

$apps = @{}
function Get-App([string]$progId) {
    if (-not $apps.ContainsKey($progId)) {
        $a = New-Object -ComObject $progId
        if ($progId -ne 'PowerPoint.Application') { $a.Visible = $false; $a.DisplayAlerts = 0 }
        $apps[$progId] = $a
    }
    $apps[$progId]
}

function Convert-Office([string]$Source, [string]$Target, $Info) {
    $app = Get-App $Info.App
    switch ($Info.App) {
        'Word.Application' {
            $doc = $app.Documents.Open($Source, $false, $true)   # ConfirmConversions, ReadOnly
            try { $doc.SaveAs2([ref]$Target, [ref]$Info.Fmt) } finally { $doc.Close([ref]0) }
        }
        'Excel.Application' {
            $wb = $app.Workbooks.Open($Source, 0, $true)
            try { $wb.SaveAs($Target, $Info.Fmt) } finally { $wb.Close($false) }
        }
        'PowerPoint.Application' {
            $pres = $app.Presentations.Open($Source, -1, 0, 0)   # ReadOnly, Untitled, WithWindow=false
            try { $pres.SaveAs($Target, $Info.Fmt) } finally { $pres.Close() }
        }
    }
}

$ok = 0; $skipped = 0; $failed = 0
$files = Get-ChildItem -LiteralPath $Path -Recurse -File |
    Where-Object { $map.ContainsKey($_.Extension.ToLower()) -and $_.Name -notlike '~$*' }

try {
    foreach ($f in $files) {
        $info = $map[$f.Extension.ToLower()]
        $target = Get-TargetPath $f $info.Ext

        if ((Test-Path -LiteralPath $target) -and -not $Overwrite) {
            Write-Verbose "Skip (exists): $target"; $skipped++; continue
        }
        if (-not $PSCmdlet.ShouldProcess($f.FullName, "Convert to $($info.Ext)")) { continue }

        try {
            if ($UseOffice) { Convert-Office $f.FullName $target $info }
            else            { Convert-OpenXml $f.FullName $target $info }
            Write-Host "OK   $($f.FullName) -> $target"
            $ok++
        }
        catch {
            Write-Warning "FAIL $($f.FullName): $($_.Exception.Message)"
            if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue }
            $failed++
        }
    }
}
finally {
    foreach ($a in $apps.Values) {
        try { $a.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "`nDone. Converted: $ok | Skipped: $skipped | Failed: $failed"
