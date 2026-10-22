#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$stage = Join-Path $env:TEMP ('batch-renamer-pack_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$destZip = Join-Path $root 'batch-renamer-1.1.0.zip'
[void](New-Item -ItemType Directory -Path $stage -Force)
Get-ChildItem -LiteralPath $root -File | Where-Object {
    $ext = $_.Extension.ToLowerInvariant()
    if ($ext -eq '.ps1' -or $ext -eq '.vbs' -or $ext -eq '.bat') { return $true }
    if ($_.Name -eq 'README.md' -or $_.Name -eq 'CHANGELOG.md') { return $true }
    if ($ext -eq '.md' -and $_.Name -notlike 'BatchRenamer_*' -and $_.Length -lt 20000) { return $true }
    return $false
} | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $stage $_.Name)
}
$copied = @(Get-ChildItem -LiteralPath $stage -File)
if ($copied.Count -lt 6) { throw "too few files: $($copied.Count)" }
if (Test-Path -LiteralPath $destZip) { Remove-Item -LiteralPath $destZip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $destZip)
Remove-Item -LiteralPath $stage -Recurse -Force
Write-Output ("ZIP=" + $destZip)
$copied | ForEach-Object { Write-Output ("FILE=" + $_.Name) }
