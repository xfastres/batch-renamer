#requires -Version 5.1
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$issPath = Join-Path $PSScriptRoot 'BatchRenamer.iss'
$distDir = Join-Path $root 'dist'
$outExe = Join-Path $distDir 'BatchRenamer-Setup-1.1.0.exe'

function Get-IsccPath {
    $candidates = @(
        'F:\xufeng\Inno Setup 6\ISCC.exe',
        'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        'C:\Program Files\Inno Setup 6\ISCC.exe'
    )
    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) { return $path }
    }

    $lnkDir = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Inno Setup 6'
    $lnk = Join-Path $lnkDir 'Inno Setup Compiler.lnk'
    if (Test-Path -LiteralPath $lnk) {
        $shell = New-Object -ComObject WScript.Shell
        $target = [string]$shell.CreateShortcut($lnk).TargetPath
        if ($target) {
            $iscc = Join-Path ([System.IO.Path]::GetDirectoryName($target)) 'ISCC.exe'
            if (Test-Path -LiteralPath $iscc) { return $iscc }
        }
    }

    $found = Get-ChildItem -Path 'C:\Program Files*','F:\' -Filter 'ISCC.exe' -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if ($found) { return $found }
    return $null
}

if (-not (Test-Path -LiteralPath $issPath)) { throw "missing $issPath" }
$iscc = Get-IsccPath
if (-not $iscc) { throw 'ISCC.exe not found. Install Inno Setup 6 first.' }

if (-not (Test-Path -LiteralPath $distDir)) {
    [void](New-Item -ItemType Directory -Path $distDir -Force)
}
if (Test-Path -LiteralPath $outExe) { Remove-Item -LiteralPath $outExe -Force }

$p = Start-Process -FilePath $iscc -ArgumentList @($issPath) -Wait -PassThru -NoNewWindow
if ($p.ExitCode -ne 0) { throw "ISCC failed, exit $($p.ExitCode)" }
if (-not (Test-Path -LiteralPath $outExe)) { throw "setup exe not created: $outExe" }

Write-Output ("ISCC=" + $iscc)
Write-Output ("SETUP=" + $outExe)
Write-Output ("SIZE=" + (Get-Item -LiteralPath $outExe).Length)
