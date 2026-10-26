<#
.SYNOPSIS
    Recursively converts .docx/.xlsx/.pptx files into .dotx/.xltx/.potx templates.
    Each template is saved in the same folder as its source file (same name, new extension).
    After a successful conversion the original is renamed by appending an extra extension
    (Report.docx -> Report.docx.bak). Nothing is deleted.

.PARAMETER Path
    Root folder to scan (recursively). Defaults to the folder the script is in.
.PARAMETER UseOffice
    Use Word/Excel/PowerPoint COM automation (SaveAs template). Requires Office installed.
    Default is a pure OpenXML approach (no Office needed): copy the file, rename the
    extension and patch the main part's content type in [Content_Types].xml.
.PARAMETER OriginalExtension
    Extension appended to the original file after conversion. Default: .bak
.PARAMETER Overwrite
    Overwrite existing templates.

.EXAMPLE
    .\Convert-ToOfficeTemplates.ps1                 # scans the script's own folder
.EXAMPLE
    .\Convert-ToOfficeTemplates.ps1 -Path C:\Docs
.EXAMPLE
    .\Convert-ToOfficeTemplates.ps1 -Path C:\Docs -UseOffice -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Path = $PSScriptRoot,
    [string]$OriginalExtension = '.bak',
    [switch]$UseOffice,
    [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Path = (Resolve-Path $Path).Path

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
Write-Host "Scanning: $Path"
$enumErrors = @()
$all = @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue -ErrorVariable enumErrors)
$files = @($all | Where-Object { $map.ContainsKey($_.Extension.ToLower()) -and $_.Name -notlike '~$*' })
Write-Host ("Found {0} files in total, {1} to convert ({2})" -f $all.Count, $files.Count,
    (($files | Group-Object { $_.Extension.ToLower() } | ForEach-Object { "$($_.Count) $($_.Name)" }) -join ', '))
foreach ($e in $enumErrors) { Write-Warning "Could not read: $($e.TargetObject) - $($e.Exception.Message)" }

try {
    foreach ($f in $files) {
        $info = $map[$f.Extension.ToLower()]
        $target = Get-TargetPath $f $info.Ext

        if ((Test-Path -LiteralPath $target) -and -not $Overwrite) {
            Write-Host "SKIP $($f.FullName) (template exists, use -Overwrite)"; $skipped++; continue
        }
        if (-not $PSCmdlet.ShouldProcess($f.FullName, "Convert to $($info.Ext)")) { continue }

        try {
            if ($UseOffice) { Convert-Office $f.FullName $target $info }
            else            { Convert-OpenXml $f.FullName $target $info }
            Write-Host "OK   $($f.FullName) -> $target"
            try { Rename-Item -LiteralPath $f.FullName -NewName ($f.Name + $OriginalExtension) }
            catch { Write-Warning "Converted, but could not rename original $($f.FullName): $($_.Exception.Message)" }
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
