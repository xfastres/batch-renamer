#requires -Version 5.1
# Phase-1 tests. ASCII-only to stay parse-safe on Windows PowerShell 5.1.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSScriptRoot
. (Join-Path $here 'BatchRenamer.Core.ps1')

$script:Pass = 0
$script:Fail = 0
$script:Skip = 0
$script:Results = New-Object System.Collections.ArrayList

function Write-Case {
    param([string]$Id, [string]$Status, [string]$Detail)
    [void]$script:Results.Add("$Id|$Status|$Detail")
    Write-Output ("{0} {1} {2}" -f $Id, $Status, $Detail)
    switch ($Status) {
        'PASS' { $script:Pass++ }
        'FAIL' { $script:Fail++ }
        default { $script:Skip++ }
    }
}

function New-PlanRow {
    param([string]$Folder, [string]$SourceName, [string]$NewName, [string]$Content = 'x')
    $path = Join-Path $Folder $SourceName
    if (-not (Test-Path -LiteralPath $path)) {
        Set-Content -LiteralPath $path -Value $Content -Encoding UTF8
    }
    return [pscustomobject]@{
        Status     = 'ke'
        SourcePath = $path
        ParentDir  = $Folder
        SourceName = $SourceName
        NewName    = $NewName
    }
}

# Status must match engine: executable status is Chinese in production.
# Rebuild helper with the real status string via here-string from unicode escape.
function Get-ExecutableStatus {
    return [string]([char]0x53EF) + [char]0x6267 + [char]0x884C
}

function New-ExecRow {
    param([string]$Folder, [string]$SourceName, [string]$NewName, [string]$Content = 'x')
    $row = New-PlanRow -Folder $Folder -SourceName $SourceName -NewName $NewName -Content $Content
    $row.Status = Get-ExecutableStatus
    return $row
}

function Assert-FileContent {
    param([string]$Path, [string]$Expected)
    if (-not (Test-Path -LiteralPath $Path)) { throw "missing file: $Path" }
    $actual = (Get-Content -LiteralPath $Path -Raw).Trim()
    if ($actual -ne $Expected) { throw "content mismatch $Path expected $Expected actual $actual" }
}

$official = Get-BatchRenamerLogFolder
$officialBefore = Get-BatchRenamerLogFingerprint -LogFolder $official
$root = Join-Path $env:TEMP ('BatchRenamerTests_' + [guid]::NewGuid().ToString('N'))
$logs = Join-Path $root 'logs'
[void](New-Item -ItemType Directory -Path $logs -Force)

try {
    try {
        $d = Join-Path $root 't01'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(New-ExecRow -Folder $d -SourceName 'a.txt' -NewName 'b.txt' -Content 'AAA')
        $r = Invoke-RenamePlan -Plan $plan -LogFolder $logs
        if ($r.SuccessCount -ne 1) { throw "success $($r.SuccessCount)" }
        Assert-FileContent -Path (Join-Path $d 'b.txt') -Expected 'AAA'
        $u = Invoke-UndoLatest -LogFolder $logs
        if (-not $u.MarkedDone) { throw 'not archived' }
        Assert-FileContent -Path (Join-Path $d 'a.txt') -Expected 'AAA'
        Write-Case -Id 'T-01' -Status 'PASS' -Detail $r.Outcome
    } catch { Write-Case -Id 'T-01' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't02'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(
            (New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'B.txt' -Content 'AAA'),
            (New-ExecRow -Folder $d -SourceName 'B.txt' -NewName 'A.txt' -Content 'BBB')
        )
        $r = Invoke-RenamePlan -Plan $plan -LogFolder $logs
        if ($r.SuccessCount -ne 2) { throw "success $($r.SuccessCount) $($r.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'BBB'
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'AAA'
        $u = Invoke-UndoLatest -LogFolder $logs
        if (-not $u.MarkedDone) { throw "undo incomplete $($u.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'AAA'
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'BBB'
        Write-Case -Id 'T-02' -Status 'PASS' -Detail 'swap+undo'
    } catch { Write-Case -Id 'T-02' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't03'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(
            (New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'B.txt' -Content 'AAA'),
            (New-ExecRow -Folder $d -SourceName 'B.txt' -NewName 'C.txt' -Content 'BBB'),
            (New-ExecRow -Folder $d -SourceName 'C.txt' -NewName 'A.txt' -Content 'CCC')
        )
        $r = Invoke-RenamePlan -Plan $plan -LogFolder $logs
        if ($r.SuccessCount -ne 3) { throw "success $($r.SuccessCount) $($r.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'CCC'
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'AAA'
        Assert-FileContent -Path (Join-Path $d 'C.txt') -Expected 'BBB'
        $u = Invoke-UndoLatest -LogFolder $logs
        if (-not $u.MarkedDone) { throw "undo incomplete $($u.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'AAA'
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'BBB'
        Assert-FileContent -Path (Join-Path $d 'C.txt') -Expected 'CCC'
        Write-Case -Id 'T-03' -Status 'PASS' -Detail 'cycle+undo'
    } catch { Write-Case -Id 'T-03' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't04'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(
            (New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'X.txt' -Content 'AAA'),
            (New-ExecRow -Folder $d -SourceName 'B.txt' -NewName 'Y.txt' -Content 'BBB')
        )
        [void](Invoke-RenamePlan -Plan $plan -LogFolder $logs)
        Set-Content -LiteralPath (Join-Path $d 'A.txt') -Value 'OCCUPY' -Encoding UTF8
        $u1 = Invoke-UndoLatest -LogFolder $logs
        if ($u1.MarkedDone) { throw 'should not archive while occupied' }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'OCCUPY'
        if (-not (Test-Path -LiteralPath (Join-Path $d 'B.txt'))) { throw 'B should be restored' }
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'BBB'
        Remove-Item -LiteralPath (Join-Path $d 'A.txt') -Force
        $u2 = Invoke-UndoLatest -LogFolder $logs
        if (-not $u2.MarkedDone) { throw "retry incomplete $($u2.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'AAA'
        Assert-FileContent -Path (Join-Path $d 'B.txt') -Expected 'BBB'
        Write-Case -Id 'T-04' -Status 'PASS' -Detail 'retry after occupy'
    } catch { Write-Case -Id 'T-04' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't05'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'Z.txt' -Content 'AAA')
        [void](Invoke-RenamePlan -Plan $plan -LogFolder $logs)
        Set-Content -LiteralPath (Join-Path $d 'A.txt') -Value 'EXTERNAL' -Encoding UTF8
        $u = Invoke-UndoLatest -LogFolder $logs
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'EXTERNAL'
        if ($u.MarkedDone) { throw 'should not archive when external occupies original' }
        if ($u.NeedsAttentionCount -lt 1) { throw 'expected NeedsAttention' }
        Write-Case -Id 'T-05' -Status 'PASS' -Detail 'no overwrite'
    } catch { Write-Case -Id 'T-05' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't06'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'N.txt' -Content 'AAA')
        [void](Invoke-RenamePlan -Plan $plan -LogFolder $logs)
        Set-Content -LiteralPath (Join-Path $d 'A.txt') -Value 'NEW' -Encoding UTF8
        $u = Invoke-UndoLatest -LogFolder $logs
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'NEW'
        Assert-FileContent -Path (Join-Path $d 'N.txt') -Expected 'AAA'
        if ($u.NeedsAttentionCount -lt 1) { throw 'expected NeedsAttention' }
        Write-Case -Id 'T-06' -Status 'PASS' -Detail 'both exist'
    } catch { Write-Case -Id 'T-06' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't07'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $badLogs = Join-Path $root 't07-notdir'
        Set-Content -LiteralPath $badLogs -Value 'file' -Encoding UTF8
        $plan = @(New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'B.txt' -Content 'AAA')
        $threw = $false
        try {
            [void](Invoke-RenamePlan -Plan $plan -LogFolder $badLogs)
        } catch { $threw = $true }
        if (-not $threw) { throw 'unwritable log folder should reject rename' }
        if (-not (Test-Path -LiteralPath (Join-Path $d 'A.txt'))) { throw 'source should remain' }
        Write-Case -Id 'T-07' -Status 'PASS' -Detail 'blocked before rename'
    } catch { Write-Case -Id 'T-07' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't09'
        [void](New-Item -ItemType Directory -Path $d -Force)
        $plan = @(
            (New-ExecRow -Folder $d -SourceName 'A.txt' -NewName 'X.txt' -Content 'AAA'),
            (New-ExecRow -Folder $d -SourceName 'B.txt' -NewName 'Y.txt' -Content 'BBB')
        )
        $hook = {
            param($Phase, $Journal)
            if ($Phase -eq 'AfterItemStaged') { throw 'TEST_INTERRUPT_STAGE1' }
        }
        $threw = $false
        try { [void](Invoke-RenamePlan -Plan $plan -LogFolder $logs -TestHook $hook) } catch {
            if ($_.Exception.Message -ne 'TEST_INTERRUPT_STAGE1') { throw }
            $threw = $true
        }
        if (-not $threw) { throw 'hook did not fire' }
        $pending = @(Get-PendingBatchRenamerJournals -LogFolder $logs | Where-Object {
            $_.Kind -eq 'Rename' -and $_.Items -and ($_.Items[0].ParentDir -eq $d)
        })
        if ($pending.Count -lt 1) { throw 'pending journal not found' }
        $temps = @(Get-ChildItem -LiteralPath $d -Filter '~rn_*')
        if ($temps.Count -lt 1) { throw 'temp name missing after interrupt' }
        $rec = Invoke-BatchRenamerRecoverJournal -Journal $pending[0] -LogFolder $logs
        if ($rec.SuccessCount -ne 2) { throw "recover success $($rec.SuccessCount) $($rec.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'X.txt') -Expected 'AAA'
        Assert-FileContent -Path (Join-Path $d 'Y.txt') -Expected 'BBB'
        Write-Case -Id 'T-09' -Status 'PASS' -Detail 'recover after stage1 interrupt'
    } catch { Write-Case -Id 'T-09' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $renameCmd = (Get-Command Invoke-RenamePlan).ScriptBlock.ToString()
        $composerCmd = (Get-Command Invoke-ComposerRename).ScriptBlock.ToString()
        if ($renameCmd -notmatch 'Invoke-BatchRenamerTransaction') { throw 'sequence executor not using core' }
        if ($composerCmd -notmatch 'Invoke-BatchRenamerTransaction') { throw 'composer executor not using core' }
        Write-Case -Id 'T-10' -Status 'PASS' -Detail 'shared transaction'
    } catch { Write-Case -Id 'T-10' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $d = Join-Path $root 't11'
        [void](New-Item -ItemType Directory -Path $d -Force)
        Set-Content -LiteralPath (Join-Path $d 'N.txt') -Value 'OLDCSV' -Encoding UTF8
        Start-Sleep -Milliseconds 20
        $csv = Join-Path $logs ('undo_{0}.csv' -f (Get-Date).ToString('yyyyMMdd_HHmmss_fff'))
        @(
            [pscustomobject]@{ OldPath = (Join-Path $d 'A.txt'); NewPath = (Join-Path $d 'N.txt'); Time = (Get-Date).ToString('s') }
        ) | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $u = Invoke-UndoLatest -LogFolder $logs
        if (-not $u.MarkedDone) { throw "legacy csv undo failed $($u.Errors -join ';')" }
        Assert-FileContent -Path (Join-Path $d 'A.txt') -Expected 'OLDCSV'
        Write-Case -Id 'T-11' -Status 'PASS' -Detail 'legacy csv'
    } catch { Write-Case -Id 'T-11' -Status 'FAIL' -Detail $_.Exception.Message }

    try {
        $officialNow = Get-BatchRenamerLogFingerprint -LogFolder $official
        if ($officialNow -ne $officialBefore) { throw 'official logs changed before cleanup' }
        Write-Case -Id 'T-12' -Status 'PASS' -Detail 'official logs unchanged'
    } catch { Write-Case -Id 'T-12' -Status 'FAIL' -Detail $_.Exception.Message }

} finally {
    $officialAfter = Get-BatchRenamerLogFingerprint -LogFolder $official
    if ($officialAfter -ne $officialBefore) {
        Write-Case -Id 'T-08a' -Status 'FAIL' -Detail 'official logs modified'
    } else {
        Write-Case -Id 'T-08a' -Status 'PASS' -Detail 'official fingerprint unchanged'
    }
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ''
Write-Output ("PASS={0} FAIL={1} SKIP={2}" -f $script:Pass, $script:Fail, $script:Skip)
if ($script:Fail -gt 0) { exit 1 }
exit 0
