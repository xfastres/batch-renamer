#requires -Version 5.1
<#
批量重命名核心：事务日志、两阶段改名、两阶段撤销、崩溃恢复。
兼容 Windows PowerShell 5.1。由主程序与测试脚本点源调用。
#>

$script:BatchRenamerLockStream = $null
$script:BatchRenamerLockPath = $null
$script:BatchRenamerDefaultLogFolder = Join-Path $env:TEMP 'BatchRenamerLogs'

function Get-BatchRenamerLogFolder {
    param([string]$LogFolder)
    if (-not [string]::IsNullOrWhiteSpace($LogFolder)) { return $LogFolder }
    if ($script:LogFolder) { return $script:LogFolder }
    return $script:BatchRenamerDefaultLogFolder
}

function Get-BatchRenamerPathKey {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return $Path.TrimEnd('\').ToLowerInvariant()
}

function Test-BatchRenamerSamePath {
    param([string]$Left, [string]$Right)
    return (Get-BatchRenamerPathKey -Path $Left) -eq (Get-BatchRenamerPathKey -Path $Right)
}

function Write-BatchRenamerAtomic {
    # 先写临时文件再替换，避免半份 JSON
    param([string]$Path, [string]$Content)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        [void](New-Item -ItemType Directory -Path $dir -Force)
    }
    $tmp = $Path + '.tmp'
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tmp, $Content, $utf8)
    if (Test-Path -LiteralPath $Path) {
        Move-Item -LiteralPath $tmp -Destination $Path -Force
    } else {
        Move-Item -LiteralPath $tmp -Destination $Path
    }
}

function Test-BatchRenamerLogFolderWritable {
    param([string]$LogFolder)
    if ([string]::IsNullOrWhiteSpace($LogFolder)) {
        throw '日志目录为空。'
    }
    if (Test-Path -LiteralPath $LogFolder -PathType Leaf) {
        throw "日志路径是文件不是目录：$LogFolder"
    }
    if (-not (Test-Path -LiteralPath $LogFolder)) {
        [void](New-Item -ItemType Directory -Path $LogFolder -Force)
    }
    $probe = Join-Path $LogFolder ('probe_{0}.tmp' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        [System.IO.File]::WriteAllText($probe, 'ok')
    } catch {
        throw "日志目录不可写：$LogFolder ；$($_.Exception.Message)"
    } finally {
        if (Test-Path -LiteralPath $probe) {
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-BatchRenamerLogFingerprint {
    param([string]$LogFolder)
    if (-not (Test-Path -LiteralPath $LogFolder)) { return 'EMPTY' }
    $files = @(Get-ChildItem -LiteralPath $LogFolder -Force -ErrorAction SilentlyContinue | Sort-Object FullName)
    if ($files.Count -eq 0) { return 'EMPTY' }
    $lines = New-Object System.Collections.ArrayList
    foreach ($file in $files) {
        $hash = ''
        if (-not $file.PSIsContainer) {
            $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
        [void]$lines.Add(('{0}|{1}|{2}|{3}' -f $file.Name, $file.Length, $file.LastWriteTimeUtc.Ticks, $hash))
    }
    return ($lines -join "`n")
}

function Enter-BatchRenamerLogLock {
    param([string]$LogFolder)
    if ($script:BatchRenamerLockStream) { return }
    Test-BatchRenamerLogFolderWritable -LogFolder $LogFolder
    $lockPath = Join-Path $LogFolder 'batchrenamer.lock'
    $maxTry = 2
    for ($i = 0; $i -lt $maxTry; $i++) {
        try {
            $script:BatchRenamerLockStream = [System.IO.File]::Open(
                $lockPath,
                [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None)
            $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$PID)
            $script:BatchRenamerLockStream.Write($bytes, 0, $bytes.Length)
            $script:BatchRenamerLockStream.Flush()
            $script:BatchRenamerLockPath = $lockPath
            return
        } catch {
            if (Test-Path -LiteralPath $lockPath) {
                $ownerPid = 0
                try {
                    $ownerPid = [int]([System.IO.File]::ReadAllText($lockPath).Trim())
                } catch { $ownerPid = 0 }
                $alive = $false
                if ($ownerPid -gt 0) {
                    try {
                        $null = Get-Process -Id $ownerPid -ErrorAction Stop
                        $alive = $true
                    } catch { $alive = $false }
                }
                if ($alive) {
                    throw "另一个批量重命名实例正在使用日志目录（PID $ownerPid）：$LogFolder"
                }
                Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
            } else {
                throw "无法创建日志锁：$($_.Exception.Message)"
            }
        }
    }
    throw "无法获取日志目录锁：$LogFolder"
}

function Exit-BatchRenamerLogLock {
    if ($script:BatchRenamerLockStream) {
        try { $script:BatchRenamerLockStream.Close() } catch { }
        try { $script:BatchRenamerLockStream.Dispose() } catch { }
        $script:BatchRenamerLockStream = $null
    }
    if ($script:BatchRenamerLockPath -and (Test-Path -LiteralPath $script:BatchRenamerLockPath)) {
        Remove-Item -LiteralPath $script:BatchRenamerLockPath -Force -ErrorAction SilentlyContinue
    }
    $script:BatchRenamerLockPath = $null
}

function Get-BatchRenamerItemSnapshot {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $item = Get-Item -LiteralPath $Path -Force
    $isDir = [bool]$item.PSIsContainer
    return [pscustomobject]@{
        Path             = $item.FullName
        IsDirectory      = $isDir
        Length           = if ($isDir) { $null } else { [int64]$item.Length }
        CreationTimeUtc  = $item.CreationTimeUtc.ToString('o')
        LastWriteTimeUtc = $item.LastWriteTimeUtc.ToString('o')
    }
}

function ConvertTo-BatchRenamerBool {
    param($Value)
    if ($Value -is [bool]) { return $Value }
    if ($null -eq $Value -or $Value -eq '') { return $false }
    return [System.Convert]::ToBoolean($Value)
}

function ConvertTo-BatchRenamerInt64 {
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return $null }
    return [int64]$Value
}

function Test-BatchRenamerSnapshotMatch {
    param([string]$Path, $Snapshot)
    if ($null -eq $Snapshot) { return $true }
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $now = Get-BatchRenamerItemSnapshot -Path $Path
    $snapDir = ConvertTo-BatchRenamerBool -Value $Snapshot.IsDirectory
    if ($now.IsDirectory -ne $snapDir) { return $false }
    if (-not $now.IsDirectory) {
        $snapLen = ConvertTo-BatchRenamerInt64 -Value $Snapshot.Length
        if ($null -ne $snapLen -and [int64]$now.Length -ne $snapLen) { return $false }
    }
    return $true
}

function Invoke-BatchRenamerTestHook {
    param($TestHook, [string]$Phase, $Journal)
    if ($null -eq $TestHook) { return }
    & $TestHook $Phase $Journal
}

function ConvertTo-BatchRenamerJournalObject {
    param($Raw)
    $items = @()
    if ($Raw.Items) { $items = @($Raw.Items) }
    return [pscustomobject]@{
        TransactionId   = [string]$Raw.TransactionId
        Kind            = [string]$Raw.Kind
        State           = [string]$Raw.State
        CreatedAt       = [string]$Raw.CreatedAt
        UpdatedAt       = [string]$Raw.UpdatedAt
        UndoCsvPath     = [string]$Raw.UndoCsvPath
        JournalPath     = [string]$Raw.JournalPath
        Error           = [string]$Raw.Error
        Items           = $items
    }
}

function Save-BatchRenamerJournal {
    param($Journal, [string]$LogFolder)
    $Journal.UpdatedAt = (Get-Date).ToString('o')
    if ([string]::IsNullOrWhiteSpace($Journal.JournalPath)) {
        $Journal.JournalPath = Join-Path $LogFolder ('txn_{0}.json' -f $Journal.TransactionId)
    }
    $json = $Journal | ConvertTo-Json -Depth 8
    Write-BatchRenamerAtomic -Path $Journal.JournalPath -Content $json
    return $Journal.JournalPath
}

function Read-BatchRenamerJournal {
    param([string]$Path)
    $text = [System.IO.File]::ReadAllText($Path)
    $raw = $text | ConvertFrom-Json
    $journal = ConvertTo-BatchRenamerJournalObject -Raw $raw
    $journal.JournalPath = $Path
    return $journal
}

function Complete-BatchRenamerJournal {
    param($Journal)
    if (-not $Journal.JournalPath -or -not (Test-Path -LiteralPath $Journal.JournalPath)) { return }
    if ($Journal.JournalPath.EndsWith('.done')) { return }
    $donePath = $Journal.JournalPath + '.done'
    Move-Item -LiteralPath $Journal.JournalPath -Destination $donePath -Force
    $Journal.JournalPath = $donePath
}

function New-BatchRenamerTempName {
    param([string]$ParentDir)
    for ($i = 0; $i -lt 8; $i++) {
        $name = '~rn_' + [guid]::NewGuid().ToString('N').Substring(0, 10)
        $path = Join-Path $ParentDir $name
        if (-not (Test-Path -LiteralPath $path)) { return $name }
    }
    throw "无法分配唯一临时名称：$ParentDir"
}

function New-BatchRenamerRenameJournal {
    param(
        [object[]]$Plan,
        [string]$LogFolder,
        [string]$Kind = 'Rename'
    )
    $id = [guid]::NewGuid().ToString('N')
    $items = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($row in @($Plan)) {
        if ($row.Status -and $row.Status -ne '可执行') { continue }
        $index++
        $parentDir = [string]$row.ParentDir
        if ([string]::IsNullOrWhiteSpace($parentDir)) {
            $parentDir = Split-Path -Parent $row.SourcePath
        }
        $originalPath = [string]$row.SourcePath
        $finalPath = Join-Path $parentDir $row.NewName
        $snap = $null
        if ($row.PSObject.Properties['SourceLength'] -and $null -ne $row.SourceLength) {
            $snap = [pscustomobject]@{
                IsDirectory      = [bool]$row.SourceIsDirectory
                Length           = $row.SourceLength
                LastWriteTimeUtc = [string]$row.SourceLastWriteTimeUtc
            }
        } elseif (Test-Path -LiteralPath $originalPath) {
            $snap = Get-BatchRenamerItemSnapshot -Path $originalPath
        }
        [void]$items.Add([pscustomobject]@{
            Id               = [string]$index
            OriginalPath     = $originalPath
            FinalPath        = $finalPath
            TempPath         = ''
            ParentDir        = $parentDir
            SourceName       = [string]$row.SourceName
            NewName          = [string]$row.NewName
            IsDirectory      = if ($snap) { [bool]$snap.IsDirectory } else { $false }
            Length           = if ($snap) { $snap.Length } else { $null }
            CreationTimeUtc  = if ($snap) { $snap.CreationTimeUtc } else { '' }
            LastWriteTimeUtc = if ($snap) { $snap.LastWriteTimeUtc } else { '' }
            State            = 'Prepared'
            Error            = ''
        })
    }
    return [pscustomobject]@{
        TransactionId = $id
        Kind          = $Kind
        State         = 'Prepared'
        CreatedAt     = (Get-Date).ToString('o')
        UpdatedAt     = (Get-Date).ToString('o')
        UndoCsvPath   = ''
        JournalPath   = Join-Path $LogFolder ('txn_{0}.json' -f $id)
        Error         = ''
        Items         = $items.ToArray()
    }
}

function Get-BatchRenamerJournalCounts {
    param($Journal)
    $items = @($Journal.Items)
    return [pscustomobject]@{
        Total            = $items.Count
        Prepared         = @($items | Where-Object { $_.State -eq 'Prepared' }).Count
        Staged           = @($items | Where-Object { $_.State -eq 'Staged' }).Count
        Applied          = @($items | Where-Object { $_.State -eq 'Applied' }).Count
        Restored         = @($items | Where-Object { $_.State -eq 'Restored' }).Count
        NeedsAttention   = @($items | Where-Object { $_.State -eq 'NeedsAttention' }).Count
    }
}

function New-BatchRenamerResult {
    param($Journal, [string[]]$Errors, [string]$Outcome)
    $counts = Get-BatchRenamerJournalCounts -Journal $Journal
    if ([string]::IsNullOrWhiteSpace($Outcome)) {
        if ($counts.NeedsAttention -gt 0) { $Outcome = 'NeedsAttention' }
        elseif ($counts.Applied -gt 0 -and ($counts.Applied + $counts.Restored) -eq $counts.Total) { $Outcome = 'Success' }
        elseif ($counts.Applied -gt 0) { $Outcome = 'PartialSuccess' }
        elseif ($counts.Restored -eq $counts.Total -and $counts.Total -gt 0) { $Outcome = 'RolledBack' }
        else { $Outcome = 'CompleteFailure' }
    }
    return [pscustomobject]@{
        TransactionId        = $Journal.TransactionId
        Outcome              = $Outcome
        SuccessCount         = $counts.Applied
        RestoredCount        = $counts.Restored
        NeedsAttentionCount  = $counts.NeedsAttention
        Errors               = @($Errors)
        LogPath              = $Journal.UndoCsvPath
        JournalPath          = $Journal.JournalPath
        MarkedDone           = ($Journal.State -eq 'Completed')
        RestoredCountUndo    = $counts.Restored
        Journal              = $Journal
    }
}

function Save-BatchRenamerUndoCsv {
    param($Journal, [string]$LogFolder)
    $applied = @($Journal.Items | Where-Object { $_.State -eq 'Applied' })
    if ($applied.Count -eq 0) { return '' }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss_fff')
    $csvPath = Join-Path $LogFolder ('undo_{0}.csv' -f $stamp)
    $records = foreach ($item in $applied) {
        [pscustomobject]@{
            OldPath         = $item.OriginalPath
            NewPath         = $item.FinalPath
            Time            = (Get-Date).ToString('s')
            TransactionId   = $Journal.TransactionId
        }
    }
    $records | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    $Journal.UndoCsvPath = $csvPath
    return $csvPath
}

function Set-BatchRenamerItemAttention {
    param($Item, [string]$Message)
    $Item.State = 'NeedsAttention'
    $Item.Error = $Message
}

function Invoke-BatchRenamerStageItems {
    param($Journal, [string]$LogFolder, $TestHook)
    $Journal.State = 'Staging'
    [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
    foreach ($item in @($Journal.Items)) {
        if ($item.State -ne 'Prepared') { continue }
        $fromPath = if ($Journal.Kind -eq 'Undo') { $item.FinalPath } else { $item.OriginalPath }
        if (-not (Test-Path -LiteralPath $fromPath)) {
            # 撤销时若原名已在、目标不在，视为已经还原
            if ($Journal.Kind -eq 'Undo' -and (Test-Path -LiteralPath $item.OriginalPath)) {
                $item.State = 'Restored'
                $item.Error = ''
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                continue
            }
            Set-BatchRenamerItemAttention -Item $item -Message "当前路径不存在：$fromPath"
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            continue
        }
        if ($Journal.Kind -ne 'Undo') {
            $snap = [pscustomobject]@{
                IsDirectory = [bool]$item.IsDirectory
                Length      = $item.Length
            }
            if (-not (Test-BatchRenamerSnapshotMatch -Path $fromPath -Snapshot $snap)) {
                Set-BatchRenamerItemAttention -Item $item -Message "源对象已变化，拒绝改名：$fromPath"
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                continue
            }
        }
        try {
            $tempName = New-BatchRenamerTempName -ParentDir $item.ParentDir
            $tempPath = Join-Path $item.ParentDir $tempName
            Rename-Item -LiteralPath $fromPath -NewName $tempName -ErrorAction Stop
            $item.TempPath = $tempPath
            $item.State = 'Staged'
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            Invoke-BatchRenamerTestHook -TestHook $TestHook -Phase 'AfterItemStaged' -Journal $Journal
        } catch {
            if ($_.Exception.Message -like 'TEST_INTERRUPT*') { throw }
            Set-BatchRenamerItemAttention -Item $item -Message "阶段1失败：$fromPath ；$($_.Exception.Message)"
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
        }
    }
}

function Invoke-BatchRenamerApplyItems {
    param($Journal, [string]$LogFolder, $TestHook)
    $Journal.State = 'Applying'
    [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
    foreach ($item in @($Journal.Items)) {
        if ($item.State -ne 'Staged') { continue }
        $toPath = if ($Journal.Kind -eq 'Undo') { $item.OriginalPath } else { $item.FinalPath }
        $toName = Split-Path -Leaf $toPath
        if (-not (Test-Path -LiteralPath $item.TempPath)) {
            Set-BatchRenamerItemAttention -Item $item -Message "临时文件丢失：$($item.TempPath)"
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            continue
        }
        if (Test-Path -LiteralPath $toPath) {
            # 目标被占用：不覆盖。撤销失败则保留临时名；改名失败则尝试退回原名
            if ($Journal.Kind -eq 'Undo') {
                Set-BatchRenamerItemAttention -Item $item -Message "原名称已被占用，未覆盖：$toPath"
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                continue
            }
            try {
                $origName = Split-Path -Leaf $item.OriginalPath
                Rename-Item -LiteralPath $item.TempPath -NewName $origName -ErrorAction Stop
                $item.State = 'Restored'
                $item.TempPath = ''
                $item.Error = "目标已存在，已退回原名：$toPath"
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            } catch {
                Set-BatchRenamerItemAttention -Item $item -Message "目标已存在且退回原名失败：$toPath ；仍为 $($item.TempPath)"
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            }
            continue
        }
        try {
            Rename-Item -LiteralPath $item.TempPath -NewName $toName -ErrorAction Stop
            if ($Journal.Kind -eq 'Undo') {
                $item.State = 'Restored'
            } else {
                $item.State = 'Applied'
            }
            $item.TempPath = ''
            $item.Error = ''
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            $hookPhase = if ($Journal.Kind -eq 'Undo') { 'AfterItemRestored' } else { 'AfterItemApplied' }
            Invoke-BatchRenamerTestHook -TestHook $TestHook -Phase $hookPhase -Journal $Journal
        } catch {
            if ($_.Exception.Message -like 'TEST_INTERRUPT*') { throw }
            if ($Journal.Kind -eq 'Undo') {
                Set-BatchRenamerItemAttention -Item $item -Message "撤销阶段2失败：$($_.Exception.Message)；仍为 $($item.TempPath)"
                [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            } else {
                try {
                    $origName = Split-Path -Leaf $item.OriginalPath
                    Rename-Item -LiteralPath $item.TempPath -NewName $origName -ErrorAction Stop
                    $item.State = 'Restored'
                    $item.TempPath = ''
                    $item.Error = "阶段2失败已退回原名：$($_.Exception.Message)"
                    [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                } catch {
                    Set-BatchRenamerItemAttention -Item $item -Message "阶段2失败且还原失败：$($_.Exception.Message)；仍为 $($item.TempPath)"
                    [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                }
            }
        }
    }
}

function Update-BatchRenamerItemsForRetry {
    # 重试时：仍停在临时名的对象按 Staged 继续；撤销中目标仍在的对象重新进入 Prepared
    param($Journal)
    foreach ($item in @($Journal.Items)) {
        if ($item.State -ne 'NeedsAttention') { continue }
        if ($item.TempPath -and (Test-Path -LiteralPath $item.TempPath)) {
            $item.State = 'Staged'
            $item.Error = ''
            continue
        }
        if ($Journal.Kind -eq 'Undo') {
            $origExists = Test-Path -LiteralPath $item.OriginalPath
            $finalExists = Test-Path -LiteralPath $item.FinalPath
            if ($finalExists -and -not $origExists) {
                $item.State = 'Prepared'
                $item.Error = ''
            } elseif ($origExists -and -not $finalExists) {
                $item.State = 'Restored'
                $item.Error = ''
            }
        }
    }
}

function Invoke-BatchRenamerRunJournal {
    param($Journal, [string]$LogFolder, $TestHook)
    Update-BatchRenamerItemsForRetry -Journal $Journal
    Invoke-BatchRenamerStageItems -Journal $Journal -LogFolder $LogFolder -TestHook $TestHook
    Invoke-BatchRenamerApplyItems -Journal $Journal -LogFolder $LogFolder -TestHook $TestHook
    $counts = Get-BatchRenamerJournalCounts -Journal $Journal
    $errors = @($Journal.Items | Where-Object { $_.Error } | ForEach-Object { '{0}: {1}' -f $_.SourceName, $_.Error })
    if ($Journal.Kind -eq 'Rename') {
        if ($counts.NeedsAttention -eq 0) {
            if ($counts.Applied -gt 0) {
                [void](Save-BatchRenamerUndoCsv -Journal $Journal -LogFolder $LogFolder)
                Invoke-BatchRenamerTestHook -TestHook $TestHook -Phase 'AfterCsv' -Journal $Journal
            }
            $Journal.State = 'Completed'
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            Complete-BatchRenamerJournal -Journal $Journal
        } else {
            $Journal.State = 'NeedsAttention'
            $Journal.Error = '存在待恢复对象，未覆盖任何外部文件。'
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
        }
    } else {
        if ($counts.NeedsAttention -eq 0 -and $counts.Prepared -eq 0 -and $counts.Staged -eq 0) {
            $Journal.State = 'Completed'
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
            if ($Journal.UndoCsvPath -and (Test-Path -LiteralPath $Journal.UndoCsvPath)) {
                $csvItem = Get-Item -LiteralPath $Journal.UndoCsvPath
                if (-not $csvItem.Name.EndsWith('.done')) {
                    Rename-Item -LiteralPath $Journal.UndoCsvPath -NewName ($csvItem.Name + '.done') -ErrorAction Stop
                    $Journal.UndoCsvPath = $Journal.UndoCsvPath + '.done'
                    [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
                }
            }
            Complete-BatchRenamerJournal -Journal $Journal
        } else {
            $Journal.State = 'NeedsAttention'
            [void](Save-BatchRenamerJournal -Journal $Journal -LogFolder $LogFolder)
        }
    }
    Invoke-BatchRenamerTestHook -TestHook $TestHook -Phase 'AfterComplete' -Journal $Journal
    return (New-BatchRenamerResult -Journal $Journal -Errors $errors)
}

function Invoke-BatchRenamerTransaction {
    param(
        [object[]]$Plan,
        [string]$LogFolder,
        $TestHook
    )
    $LogFolder = Get-BatchRenamerLogFolder -LogFolder $LogFolder
    $executable = @($Plan | Where-Object { $_.Status -eq '可执行' })
    if ($executable.Count -eq 0) {
        return [pscustomobject]@{
            TransactionId       = ''
            Outcome             = 'CompleteFailure'
            SuccessCount        = 0
            RestoredCount       = 0
            NeedsAttentionCount = 0
            Errors              = @('没有可执行的对象。')
            LogPath             = $null
            JournalPath         = $null
            MarkedDone          = $false
            Journal             = $null
        }
    }
    Enter-BatchRenamerLogLock -LogFolder $LogFolder
    try {
        $journal = New-BatchRenamerRenameJournal -Plan $executable -LogFolder $LogFolder -Kind Rename
        [void](Save-BatchRenamerJournal -Journal $journal -LogFolder $LogFolder)
        Invoke-BatchRenamerTestHook -TestHook $TestHook -Phase 'AfterPrepare' -Journal $journal
        foreach ($item in @($journal.Items)) {
            if (-not (Test-Path -LiteralPath $item.OriginalPath)) {
                Remove-Item -LiteralPath $journal.JournalPath -Force -ErrorAction SilentlyContinue
                throw "源对象已不存在，请重新预览：$($item.SourceName)"
            }
            $snap = [pscustomobject]@{
                IsDirectory = $item.IsDirectory
                Length      = $item.Length
            }
            if (-not (Test-BatchRenamerSnapshotMatch -Path $item.OriginalPath -Snapshot $snap)) {
                Remove-Item -LiteralPath $journal.JournalPath -Force -ErrorAction SilentlyContinue
                throw "源对象已变化，请重新预览：$($item.SourceName)"
            }
        }
        return Invoke-BatchRenamerRunJournal -Journal $journal -LogFolder $LogFolder -TestHook $TestHook
    } finally {
        Exit-BatchRenamerLogLock
    }
}

function Get-UndoLogs {
    param([string]$LogFolder)
    $LogFolder = Get-BatchRenamerLogFolder -LogFolder $LogFolder
    if (-not (Test-Path -LiteralPath $LogFolder)) { return @() }
    return @(Get-ChildItem -LiteralPath $LogFolder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'undo_*.csv' -and $_.Name -notlike '*.done' -and $_.Name -notlike '*.txn.json' } |
        Sort-Object Name -Descending)
}

function Get-UndoCount {
    param([string]$LogFolder)
    return @(Get-UndoLogs -LogFolder $LogFolder).Count
}

function Get-PendingBatchRenamerJournals {
    param([string]$LogFolder)
    $LogFolder = Get-BatchRenamerLogFolder -LogFolder $LogFolder
    if (-not (Test-Path -LiteralPath $LogFolder)) { return @() }
    $list = New-Object System.Collections.ArrayList
    $files = @(Get-ChildItem -LiteralPath $LogFolder -File -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.Name -like 'txn_*.json' -or $_.Name -like '*.csv.txn.json') -and
            $_.Name -notlike '*.done' -and $_.Name -notlike '*.tmp'
        })
    foreach ($file in $files) {
        try {
            [void]$list.Add((Read-BatchRenamerJournal -Path $file.FullName))
        } catch { }
    }
    return $list.ToArray()
}

function New-BatchRenamerUndoJournalFromCsv {
    param([string]$CsvPath, [string]$LogFolder)
    $records = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8)
    $id = [guid]::NewGuid().ToString('N')
    $batchFinalKeys = @{}
    foreach ($record in $records) {
        $batchFinalKeys[(Get-BatchRenamerPathKey -Path $record.NewPath)] = $true
    }
    $items = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($record in $records) {
        $index++
        $originalPath = [string]$record.OldPath
        $finalPath = [string]$record.NewPath
        $parentDir = Split-Path -Parent $finalPath
        if ([string]::IsNullOrWhiteSpace($parentDir)) { $parentDir = Split-Path -Parent $originalPath }
        $state = 'Prepared'
        $errorText = ''
        $origExists = Test-Path -LiteralPath $originalPath
        $finalExists = Test-Path -LiteralPath $finalPath
        $origInBatchFinal = $batchFinalKeys.ContainsKey((Get-BatchRenamerPathKey -Path $originalPath))
        if (-not $finalExists -and $origExists) {
            $state = 'Restored'
        } elseif ($finalExists -and $origExists -and -not $origInBatchFinal) {
            # 原路径被批外对象占用，不能当成交换场景
            $state = 'NeedsAttention'
            $errorText = "原路径被批外对象占用，未覆盖：$originalPath"
        } elseif (-not $finalExists -and -not $origExists) {
            $state = 'NeedsAttention'
            $errorText = "原路径与新路径都不存在：$originalPath"
        }
        $snap = $null
        $from = if ($finalExists) { $finalPath } else { $originalPath }
        if (Test-Path -LiteralPath $from) { $snap = Get-BatchRenamerItemSnapshot -Path $from }
        [void]$items.Add([pscustomobject]@{
            Id               = [string]$index
            OriginalPath     = $originalPath
            FinalPath        = $finalPath
            TempPath         = ''
            ParentDir        = $parentDir
            SourceName       = Split-Path -Leaf $originalPath
            NewName          = Split-Path -Leaf $finalPath
            IsDirectory      = if ($snap) { [bool]$snap.IsDirectory } else { $false }
            Length           = if ($snap) { $snap.Length } else { $null }
            CreationTimeUtc  = if ($snap) { $snap.CreationTimeUtc } else { '' }
            LastWriteTimeUtc = if ($snap) { $snap.LastWriteTimeUtc } else { '' }
            State            = $state
            Error            = $errorText
        })
    }
    $sidecar = $CsvPath + '.txn.json'
    return [pscustomobject]@{
        TransactionId = $id
        Kind          = 'Undo'
        State         = 'Prepared'
        CreatedAt     = (Get-Date).ToString('o')
        UpdatedAt     = (Get-Date).ToString('o')
        UndoCsvPath   = $CsvPath
        JournalPath   = $sidecar
        Error         = ''
        Items         = $items.ToArray()
    }
}

function Invoke-UndoLatest {
    param(
        [string]$LogFolder,
        $TestHook
    )
    $LogFolder = Get-BatchRenamerLogFolder -LogFolder $LogFolder
    $logs = @(Get-UndoLogs -LogFolder $LogFolder)
    if ($logs.Count -eq 0) { throw '没有找到可撤销的操作记录。' }
    $logFile = $logs[0]
    Enter-BatchRenamerLogLock -LogFolder $LogFolder
    try {
        $sidecar = $logFile.FullName + '.txn.json'
        if (Test-Path -LiteralPath $sidecar) {
            $journal = Read-BatchRenamerJournal -Path $sidecar
        } else {
            $journal = New-BatchRenamerUndoJournalFromCsv -CsvPath $logFile.FullName -LogFolder $LogFolder
            [void](Save-BatchRenamerJournal -Journal $journal -LogFolder $LogFolder)
        }
        $result = Invoke-BatchRenamerRunJournal -Journal $journal -LogFolder $LogFolder -TestHook $TestHook
        return [pscustomobject]@{
            RestoredCount = @($journal.Items | Where-Object { $_.State -eq 'Restored' }).Count
            Errors        = @($result.Errors)
            LogPath       = $logFile.FullName
            MarkedDone    = ($journal.State -eq 'Completed')
            Outcome       = $result.Outcome
            TransactionId = $journal.TransactionId
            JournalPath   = $journal.JournalPath
            NeedsAttentionCount = $result.NeedsAttentionCount
        }
    } finally {
        Exit-BatchRenamerLogLock
    }
}

function Invoke-BatchRenamerRecoverJournal {
    param($Journal, [string]$LogFolder, $TestHook)
    $LogFolder = Get-BatchRenamerLogFolder -LogFolder $LogFolder
    Enter-BatchRenamerLogLock -LogFolder $LogFolder
    try {
        return Invoke-BatchRenamerRunJournal -Journal $Journal -LogFolder $LogFolder -TestHook $TestHook
    } finally {
        Exit-BatchRenamerLogLock
    }
}

function Invoke-RenamePlan {
    param(
        [object[]]$Plan,
        [string]$LogFolder,
        $TestHook
    )
    Invoke-BatchRenamerTransaction -Plan $Plan -LogFolder $LogFolder -TestHook $TestHook
}

function Invoke-ComposerRename {
    param(
        [object[]]$Plan,
        [string]$LogFolder,
        $TestHook
    )
    Invoke-BatchRenamerTransaction -Plan $Plan -LogFolder $LogFolder -TestHook $TestHook
}

function Format-BatchRenamerPendingSummary {
    param($Journals)
    $lines = New-Object System.Collections.ArrayList
    foreach ($journal in @($Journals)) {
        $counts = Get-BatchRenamerJournalCounts -Journal $journal
        [void]$lines.Add(('事务 {0}  类型={1}  状态={2}  待处理={3}  已完成={4}  需人工={5}' -f `
            $journal.TransactionId, $journal.Kind, $journal.State, `
            ($counts.Prepared + $counts.Staged), $counts.Applied, $counts.NeedsAttention))
    }
    return ($lines -join "`r`n")
}
