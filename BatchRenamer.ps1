#requires -Version 5.1
<#
批量重命名工具（序号编号）
- 支持序号命名、仅加前缀、仅加后缀、前缀+序号+原名 四种模式
- 支持名称/修改时间/创建时间/大小/随机排序，可递归子目录、按扩展名包含或排除
- 支持查找替换（可选区分大小写）
- 预览后再执行，执行前做冲突检测，采用两阶段重命名避免互相覆盖
- 每次执行写入撤销日志，支持连续多步撤销
#>
param([string]$Folder = '', [switch]$SelfTest, [switch]$Diagnose)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Plan = @()
$script:LogFolder = Join-Path $env:TEMP 'BatchRenamerLogs'
$script:ScriptDir = $PSScriptRoot
$script:BlockedStatuses = @('无效字符', '名称重复', '目标已存在', '名称为空', '有空段')
$script:CorePath = Join-Path $PSScriptRoot 'BatchRenamer.Core.ps1'
if (-not (Test-Path -LiteralPath $script:CorePath)) {
    throw '找不到 BatchRenamer.Core.ps1，请与主程序放在同一文件夹。'
}
. $script:CorePath

# ---------------------------------------------------------------------------
# 纯逻辑函数（不依赖界面，便于自测）
# ---------------------------------------------------------------------------

function Get-NaturalKey {
    # 思路：把名称里的数字段按 20 位左补零，使 "2" 排在 "10" 前面
    param([string]$Name)
    $pieces = [regex]::Split($Name.ToLower(), '(\d+)')
    $keys = foreach ($piece in $pieces) {
        if ($piece -match '^\d+$') { $piece.PadLeft(20, '0') } else { $piece }
    }
    return ($keys -join '')
}

function Get-InvalidNameChars {
    # 文件名非法字符 + 冒号等，避免生成 Windows 不允许的名称
    $chars = [System.IO.Path]::GetInvalidFileNameChars()
    return $chars
}

function Test-NameValid {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    foreach ($ch in (Get-InvalidNameChars)) {
        if ($Name.Contains([string]$ch)) { return $false }
    }
    if ($Name.EndsWith('.') -or $Name.EndsWith(' ')) { return $false }
    $reserved = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5',
                  'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4',
                  'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    $basePart = $Name.Split('.')[0].ToUpperInvariant()
    if ($reserved -contains $basePart) { return $false }
    return $true
}

function Get-ExtensionList {
    # 思路：把 "jpg,png" 或 "*.jpg;*.png" 统一成 ".jpg/.png" 字符串数组，便于用 -contains 精确比对
    param([string]$Filter)
    if ([string]::IsNullOrWhiteSpace($Filter)) { return @() }
    $list = New-Object System.Collections.ArrayList
    foreach ($token in ($Filter -split '[,;\s]+')) {
        $clean = $token.Trim().TrimStart('*').TrimStart('.')
        if ($clean) { [void]$list.Add('.' + $clean.ToLowerInvariant()) }
    }
    return $list.ToArray()
}

function Get-ReplacedName {
    # 思路：按纯文本替换，默认不区分大小写；把替换串里的 $ 转义，避免被当作正则替换变量
    param([string]$Name, [string]$From, [string]$To, [switch]$CaseSensitive)
    if ([string]::IsNullOrEmpty($From)) { return $Name }
    $options = if ($CaseSensitive) { [System.Text.RegularExpressions.RegexOptions]::None }
               else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    $replacement = $To.Replace('$', '$$')
    return [regex]::Replace($Name, [regex]::Escape($From), $replacement, $options)
}

function Build-Plan {
    param(
        [string]$Folder,
        [ValidateSet('Files', 'Folders', 'Both')][string]$Mode = 'Files',
        [ValidateSet('Numbered', 'PrefixOnly', 'SuffixOnly', 'PrefixNumberOriginal')][string]$NameMode = 'Numbered',
        [string]$Prefix = '',
        [string]$Suffix = '',
        [int]$StartNumber = 1,
        [int]$Digits = 3,
        [ValidateSet('Name', 'LastWriteTime', 'CreationTime', 'Size', 'Random')][string]$SortBy = 'Name',
        [switch]$Recurse,
        [string]$ExtensionFilter = '',
        [ValidateSet('Include', 'Exclude')][string]$FilterMode = 'Include',
        [string]$ReplaceFrom = '',
        [string]$ReplaceTo = '',
        [switch]$CaseSensitive
    )

    $items = @()
    if ($Recurse) {
        # 思路：递归时只收集文件。若同时重命名文件夹，子树路径会在遍历后失效，
        # 因此递归模式必须忽略「文件夹」范围。
        if ($Mode -ne 'Files') { throw '递归模式下只能处理文件，请把「处理对象」改为「文件」。' }
        $items = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction Stop)
    } else {
        if ($Mode -eq 'Files' -or $Mode -eq 'Both') {
            $items += @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop)
        }
        if ($Mode -eq 'Folders' -or $Mode -eq 'Both') {
            $items += @(Get-ChildItem -LiteralPath $Folder -Directory -ErrorAction Stop)
        }
    }

    # 扩展名过滤只作用于文件；文件夹始终保留，避免递归半途丢失整棵子树
    $extensionList = @(Get-ExtensionList -Filter $ExtensionFilter)
    if ($extensionList.Count -gt 0) {
        $items = @($items | Where-Object {
            if ($_.PSIsContainer) { $true }
            else {
                $ext = [System.IO.Path]::GetExtension($_.Name).ToLowerInvariant()
                $matched = $extensionList -contains $ext
                if ($FilterMode -eq 'Exclude') { -not $matched } else { $matched }
            }
        })
    }
    if ($items.Count -eq 0) { throw '没有符合条件的对象（请检查处理范围、扩展名过滤或查找替换）。' }

    switch ($SortBy) {
        'LastWriteTime' { $sorted = @($items | Sort-Object -Property LastWriteTime, Name) }
        'CreationTime'  { $sorted = @($items | Sort-Object -Property CreationTime, Name) }
        'Size'          { $sorted = @($items | Sort-Object -Property @{ Expression = { if ($_.PSIsContainer) { 0 } else { $_.Length } } }, Name) }
        'Random'        { $sorted = @($items | Sort-Object -Property @{ Expression = { [guid]::NewGuid().ToString() } }) }
        default {
            # 先按所在目录分组，再按自然序，保证递归时各子目录内序号连续
            $sorted = @($items | Sort-Object -Property @{ Expression = { $_.DirectoryName } },
                                                   @{ Expression = { Get-NaturalKey $_.Name } })
        }
    }

    # 第一遍：先算出所有新名称
    $planItems = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($item in $sorted) {
        $index++
        $numberText = ($StartNumber + $index - 1).ToString().PadLeft($Digits, '0')
        $isDirectory = [bool]$item.PSIsContainer
        $extension = if ($isDirectory) { '' } else { [System.IO.Path]::GetExtension($item.Name) }
        $baseName = if ($isDirectory) { $item.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($item.Name) }
        $baseName = Get-ReplacedName -Name $baseName -From $ReplaceFrom -To $ReplaceTo -CaseSensitive:$CaseSensitive

        $coreName = switch ($NameMode) {
            'PrefixOnly'          { $Prefix + $baseName }
            'SuffixOnly'          { $baseName + $Suffix }
            'PrefixNumberOriginal' { $Prefix + $numberText + $Suffix + $baseName }
            default               { $Prefix + $numberText + $Suffix }
        }
        $parentDir = Split-Path -Parent $item.FullName
        # 用「路径段数」计算相对位置，避免短路径(如 YANGHU~1)与长路径前缀不一致导致截取错位
        $folderSegments = @($Folder.TrimEnd('\') -split '\\')
        $parentSegments = @($parentDir -split '\\')
        $location = if ($parentSegments.Count -gt $folderSegments.Count) {
            ($parentSegments[$folderSegments.Count..($parentSegments.Count - 1)]) -join '\'
        } else { '.' }

        [void]$planItems.Add([pscustomobject]@{
            Index                = $index
            Type                 = if ($isDirectory) { '文件夹' } else { '文件' }
            SourceName           = $item.Name
            SourcePath           = $item.FullName
            ParentDir            = $parentDir
            Location             = $location
            NewName              = $coreName + $extension
            Status               = '可执行'
            SourceIsDirectory    = $isDirectory
            SourceLength         = if ($isDirectory) { $null } else { [int64]$item.Length }
            SourceLastWriteTimeUtc = $item.LastWriteTimeUtc.ToString('o')
        })
    }

    $sourceMap = @{}
    foreach ($row in $planItems) { $sourceMap[$row.SourcePath.ToLowerInvariant()] = $row }

    # 冲突按「同目录 + 同名」统计，递归时不同子目录可以重名
    $targetCount = @{}
    foreach ($row in $planItems) {
        $key = ($row.ParentDir + '|' + $row.NewName).ToLowerInvariant()
        if ($targetCount.ContainsKey($key)) { $targetCount[$key]++ } else { $targetCount[$key] = 1 }
    }

    # 第二遍：判定状态
    foreach ($row in $planItems) {
        if ([string]::IsNullOrWhiteSpace($row.NewName)) { $row.Status = '名称为空'; continue }
        if (-not (Test-NameValid -Name $row.NewName)) { $row.Status = '无效字符'; continue }
        if ([string]::Equals($row.NewName, $row.SourceName, [System.StringComparison]::OrdinalIgnoreCase)) { $row.Status = '未变化'; continue }
        $key = ($row.ParentDir + '|' + $row.NewName).ToLowerInvariant()
        if ($targetCount[$key] -gt 1) { $row.Status = '名称重复'; continue }

        $targetPath = Join-Path $row.ParentDir $row.NewName
        if (Test-Path -LiteralPath $targetPath) {
            $owner = $sourceMap[$targetPath.ToLowerInvariant()]
            # 被本批次中「不会改名」的对象占用才算冲突；会被移走的对象可安全让出名字
            if (-not $owner -or $owner.NewName -eq $owner.SourceName) { $row.Status = '目标已存在' }
        }
    }
    return $planItems.ToArray()
}

function Format-RenameResultMessage {
    param($Result, [string]$Action = '执行')
    $outcome = [string]$Result.Outcome
    $message = "$Action 结果：$outcome`r`n成功 $($Result.SuccessCount) 个"
    if ($Result.PSObject.Properties['RestoredCount']) {
        $message += "，已退回原名 $($Result.RestoredCount) 个"
    }
    if ($Result.NeedsAttentionCount -gt 0) {
        $message += "，待恢复 $($Result.NeedsAttentionCount) 个"
    }
    $message += "，失败 $($Result.Errors.Count) 个。"
    if ($Result.TransactionId) { $message += "`r`n事务：$($Result.TransactionId)" }
    if ($Result.Errors.Count -gt 0) {
        $message += "`r`n`r`n前 10 条：`r`n" + (($Result.Errors | Select-Object -First 10) -join "`r`n")
    }
    if ($Result.LogPath) { $message += "`r`n`r`n撤销日志：$($Result.LogPath)" }
    if ($Result.JournalPath) { $message += "`r`n事务记录：$($Result.JournalPath)" }
    if ($Result.NeedsAttentionCount -gt 0) {
        $message += "`r`n`r`n请勿手动删除 ~rn_ 临时名。可稍后再次撤销或启动时选择恢复。"
    }
    return $message
}

function ConvertTo-PreviewTable {
    param([object[]]$Plan)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('序号', [int])
    [void]$table.Columns.Add('类型', [string])
    [void]$table.Columns.Add('位置', [string])
    [void]$table.Columns.Add('原名称', [string])
    [void]$table.Columns.Add('新名称', [string])
    [void]$table.Columns.Add('状态', [string])
    foreach ($row in $Plan) {
        $dataRow = $table.NewRow()
        $dataRow['序号'] = $row.Index
        $dataRow['类型'] = $row.Type
        $dataRow['位置'] = $row.Location
        $dataRow['原名称'] = $row.SourceName
        $dataRow['新名称'] = $row.NewName
        $dataRow['状态'] = $row.Status
        $table.Rows.Add($dataRow)
    }
    # 注意：直接 return DataTable 会被 PowerShell 展开成 DataRow 集合，
    # 导致 DataGridView 绑定后列和行都渲染不出来（预览空白）
    Write-Output -NoEnumerate $table
}

# ---------------------------------------------------------------------------
# 自测：在临时目录跑一遍 预览 -> 重命名 -> 撤销
# ---------------------------------------------------------------------------
function Invoke-SelfTest {
    param([string]$LogFolder)
    if ([string]::IsNullOrWhiteSpace($LogFolder)) { throw '自测必须提供独立日志目录。' }
    $rootA = Join-Path $LogFolder 'dataA'
    $rootB = Join-Path $LogFolder 'dataB'
    [void](New-Item -ItemType Directory -Path $rootA -Force)
    [void](New-Item -ItemType Directory -Path $rootB -Force)
    try {
        # 场景A：顶层 5 文件 + 2 文件夹，序号命名
        foreach ($name in @('b.txt', '10.txt', 'a.txt', '2.txt', 'c.log')) {
            Set-Content -LiteralPath (Join-Path $rootA $name) -Value 'x' -Encoding UTF8
        }
        [void](New-Item -ItemType Directory -Path (Join-Path $rootA 'dir2'))
        [void](New-Item -ItemType Directory -Path (Join-Path $rootA 'dir1'))

        $planA = Build-Plan -Folder $rootA -Mode Both -Prefix 'X_' -Suffix '' -StartNumber 1 -Digits 2 -SortBy Name
        Write-Output ("PLAN_A_COUNT={0}" -f $planA.Count)
        $planA | ForEach-Object { Write-Output ("  {0} | {1} -> {2} | {3}" -f $_.Index, $_.SourceName, $_.NewName, $_.Status) }
        $blockedA = @($planA | Where-Object { $script:BlockedStatuses -contains $_.Status })
        if ($blockedA.Count -ne 0) { throw "自测失败：方案A存在被阻止的记录 $($blockedA.Count) 条" }
        if ($planA.Count -ne 7) { throw "自测失败：期望 7 个对象，实际 $($planA.Count)" }

        # 回归检查：预览表必须是 DataTable 且行数完整，防止再次被 PowerShell 展开成 DataRow
        $table = ConvertTo-PreviewTable -Plan $planA
        Write-Output ("TABLE_TYPE={0} ROWS={1}" -f $table.GetType().Name, $table.Rows.Count)
        if ($table -isnot [System.Data.DataTable]) { throw '自测失败：预览表类型不是 DataTable' }
        if ($table.Rows.Count -ne 7) { throw '自测失败：预览表行数不符' }

        # 回归检查：真正绑定到 DataGridView 后必须有数据行，避免出现「预览空白」
        $testGrid = New-Object System.Windows.Forms.DataGridView
        $testGrid.AllowUserToAddRows = $false
        $holder = New-Object System.Windows.Forms.Form
        $holder.Controls.Add($testGrid)
        $testGrid.DataSource = $table
        Write-Output ("GRID_ROWS={0} GRID_COLUMNS={1}" -f $testGrid.Rows.Count, $testGrid.Columns.Count)
        if ($testGrid.Rows.Count -ne 7 -or $testGrid.Columns.Count -ne 6) {
            throw '自测失败：DataGridView 绑定后没有数据（预览会显示空白）'
        }
        $holder.Dispose()

        # 场景B：递归 + 扩展名过滤(jpg) + 查找替换(photo->pic) + 仅加前缀
        [void](New-Item -ItemType Directory -Path (Join-Path $rootB 'a'))
        [void](New-Item -ItemType Directory -Path (Join-Path $rootB 'b'))
        foreach ($relative in @('a\photo1.JPG', 'a\photo2.jpg', 'a\photo3.png', 'b\note.txt', 'note.txt')) {
            Set-Content -LiteralPath (Join-Path $rootB $relative) -Value 'x' -Encoding UTF8
        }
        $planB = Build-Plan -Folder $rootB -Mode Files -NameMode PrefixOnly -Prefix 'P_' `
            -Recurse -ExtensionFilter 'jpg' -ReplaceFrom 'photo' -ReplaceTo 'pic'
        Write-Output ("PLAN_B_COUNT={0}" -f $planB.Count)
        $planB | ForEach-Object { Write-Output ("  {0} | {1} -> {2} | {3}" -f $_.Index, $_.Location, $_.NewName, $_.Status) }
        if ($planB.Count -ne 2) { throw "自测失败：递归+过滤期望 2 个，实际 $($planB.Count)" }
        if (@($planB | Where-Object { $_.NewName -notmatch '^P_pic\d\.jpg$' }).Count -ne 0) {
            throw '自测失败：仅加前缀/查找替换结果不符'
        }
        if (@($planB | Where-Object { $_.Location -ne 'a' }).Count -ne 0) {
            throw '自测失败：递归位置列不符'
        }

        # 场景C：仅加后缀（不执行，只验证命名）
        $planC = Build-Plan -Folder $rootB -Mode Files -NameMode SuffixOnly -Suffix '_旧'
        $suffixRow = @($planC | Where-Object { $_.SourceName -eq 'note.txt' })
        if ($suffixRow.Count -ne 1 -or $suffixRow[0].NewName -ne 'note_旧.txt') {
            throw "自测失败：仅加后缀结果不符（$($suffixRow[0].NewName)）"
        }

        # 执行 A、B，连续撤销两次，验证多步撤销
        $resultA = Invoke-RenamePlan -Plan $planA -LogFolder $LogFolder
        Write-Output ("RENAMED_A={0} ERRORS={1} OUTCOME={2}" -f $resultA.SuccessCount, $resultA.Errors.Count, $resultA.Outcome)
        if ($resultA.SuccessCount -ne 7) { throw '自测失败：A 重命名数量不符' }
        $resultB = Invoke-RenamePlan -Plan $planB -LogFolder $LogFolder
        Write-Output ("RENAMED_B={0} ERRORS={1} OUTCOME={2}" -f $resultB.SuccessCount, $resultB.Errors.Count, $resultB.Outcome)
        if ($resultB.SuccessCount -ne 2) { throw '自测失败：B 重命名数量不符' }
        if (-not (Test-Path -LiteralPath (Join-Path $rootB 'a\P_pic1.JPG'))) { throw '自测失败：缺少 a\P_pic1.JPG' }

        $undo1 = Invoke-UndoLatest -LogFolder $LogFolder
        Write-Output ("UNDO1={0} ERRORS={1} DONE={2}" -f $undo1.RestoredCount, $undo1.Errors.Count, $undo1.MarkedDone)
        if ($undo1.RestoredCount -ne 2 -or -not $undo1.MarkedDone) { throw '自测失败：第一步撤销不符' }
        $undo2 = Invoke-UndoLatest -LogFolder $LogFolder
        Write-Output ("UNDO2={0} ERRORS={1} DONE={2}" -f $undo2.RestoredCount, $undo2.Errors.Count, $undo2.MarkedDone)
        if ($undo2.RestoredCount -ne 7 -or -not $undo2.MarkedDone) { throw '自测失败：第二步撤销不符' }

        foreach ($name in @('b.txt', '10.txt', 'a.txt', '2.txt', 'c.log', 'dir2', 'dir1')) {
            if (-not (Test-Path -LiteralPath (Join-Path $rootA $name))) { throw "自测失败：未还原 $name" }
        }
        if (-not (Test-Path -LiteralPath (Join-Path $rootB 'a\photo1.JPG'))) { throw '自测失败：未还原 a\photo1.JPG' }
        if ((Get-UndoCount -LogFolder $LogFolder) -ne 0) { throw '自测失败：测试日志撤销后仍有未完成 CSV' }
    } finally {
        Remove-Item -LiteralPath $rootA -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $rootB -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 自测入口放在所有函数定义之后（见文件末尾的界面段落）
# ---------------------------------------------------------------------------
# 分段拼接：纯逻辑（与主流程共用命名/冲突/两阶段改名规则）
# ---------------------------------------------------------------------------

function Split-NameToSegments {
    # 思路：按连接符拆分，段数不足补空、超出则多余部分并入最后一段，避免丢内容
    param(
        [string]$BaseName,
        [string]$Separator,
        [int]$SegmentCount
    )
    if ([string]::IsNullOrEmpty($Separator)) {
        $parts = @($BaseName)
    } else {
        $parts = @($BaseName -split [regex]::Escape($Separator))
    }
    $result = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $SegmentCount; $index++) {
        if ($index -lt $parts.Count - 1 -and $index -lt $SegmentCount - 1) {
            [void]$result.Add($parts[$index].Trim())
        } elseif ($index -ge $parts.Count) {
            [void]$result.Add('')
        } else {
            # 最后一段吃掉剩余部分，避免分隔符多于段数时丢内容；先逐段去空白再拼回
            $tail = @($parts[$index..($parts.Count - 1)] | ForEach-Object { $_.Trim() })
            [void]$result.Add(($tail -join $Separator).Trim())
        }
    }
    Write-Output -NoEnumerate ([string[]]$result.ToArray())
}

function Compose-Name {
    param(
        [string[]]$Segments,
        [string]$Separator,
        [string]$Extension
    )
    $joined = ($Segments -join $Separator).Trim()
    if ([string]::IsNullOrEmpty($joined)) { return '' }
    return $joined + $Extension
}

function Build-ComposerPlan {
    # 冲突规则与序号模式一致：同目录重名、批内让出名称可执行、批外占用才阻止
    param(
        [object[]]$Rows,
        [string]$Separator
    )
    $plan = New-Object System.Collections.ArrayList
    foreach ($row in $Rows) {
        $segments = [string[]]$row.Segments
        $newName = Compose-Name -Segments $segments -Separator $Separator -Extension $row.Extension
        $status = '可执行'
        $emptySegment = @($segments | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count
        if ($emptySegment -gt 0) {
            $status = '有空段'
        } elseif ([string]::IsNullOrWhiteSpace($newName)) {
            $status = '名称为空'
        } elseif (-not (Test-NameValid -Name $newName)) {
            $status = '无效字符'
        } elseif ([string]::Equals($row.Name, $newName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $status = '未变化'
        }

        [void]$plan.Add([pscustomobject]@{
            SourcePath             = $row.SourcePath
            ParentDir              = $row.ParentDir
            SourceName             = $row.Name
            NewName                = $newName
            Status                 = $status
            Segments               = $segments
            SourceIsDirectory      = $false
            SourceLength           = $row.SourceLength
            SourceLastWriteTimeUtc = $row.SourceLastWriteTimeUtc
        })
    }

    $sourceMap = @{}
    foreach ($row in $plan) { $sourceMap[(Get-BatchRenamerPathKey -Path $row.SourcePath)] = $row }
    $targetCount = @{}
    foreach ($row in $plan) {
        $key = (Get-BatchRenamerPathKey -Path ($row.ParentDir + '|' + $row.NewName))
        if ($targetCount.ContainsKey($key)) { $targetCount[$key]++ } else { $targetCount[$key] = 1 }
    }
    foreach ($row in $plan) {
        if ($script:BlockedStatuses -contains $row.Status -or $row.Status -eq '未变化' -or $row.Status -eq '有空段') { continue }
        $key = (Get-BatchRenamerPathKey -Path ($row.ParentDir + '|' + $row.NewName))
        if ($targetCount[$key] -gt 1) { $row.Status = '名称重复'; continue }
        $targetPath = Join-Path $row.ParentDir $row.NewName
        if (Test-Path -LiteralPath $targetPath) {
            $owner = $sourceMap[(Get-BatchRenamerPathKey -Path $targetPath)]
            if (-not $owner -or [string]::Equals($owner.NewName, $owner.SourceName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $row.Status = '目标已存在'
            }
        }
    }
    return $plan.ToArray()
}

# Invoke-ComposerRename 已由 BatchRenamer.Core.ps1 提供，与序号模式共用同一事务引擎。

function Get-RowRecords {
    # 思路：把界面表格转成纯数据，供预览计算与执行共用
    param($Grid, [int]$SegmentCount, [string]$FolderPath)

    $rows = New-Object System.Collections.ArrayList
    foreach ($gridRow in $Grid.Rows) {
        $name = [string]$gridRow.Cells[0].Value
        $segments = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $SegmentCount; $index++) {
            [void]$segments.Add([string]$gridRow.Cells[$index + 1].Value)
        }
        $extension = [System.IO.Path]::GetExtension($name)
        $sourcePath = Join-Path $FolderPath $name
        $length = $null
        $writeTime = ''
        if (Test-Path -LiteralPath $sourcePath -PathType Leaf) {
            $item = Get-Item -LiteralPath $sourcePath -Force
            $length = [int64]$item.Length
            $writeTime = $item.LastWriteTimeUtc.ToString('o')
        }
        [void]$rows.Add([pscustomobject]@{
            SourcePath             = $sourcePath
            ParentDir              = $FolderPath
            Name                   = $name
            Extension              = $extension
            Segments               = [string[]]$segments.ToArray()
            SourceLength           = $length
            SourceLastWriteTimeUtc = $writeTime
        })
    }
    return $rows.ToArray()
}

function Get-SegmentFolderItems {
    # 思路：只取当前目录的普通文件，排除隐藏/系统文件，按自然序排列
    param([string]$FolderPath)
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { throw '请先选择目标文件夹。' }
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { throw '目标文件夹不存在，请重新选择。' }

    $files = @(Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction Stop |
        Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::Hidden) -and
                        -not ($_.Attributes -band [System.IO.FileAttributes]::System) })
    return @($files | Sort-Object { Get-NaturalKey -Name $_.Name })
}

function Invoke-ComposerSelfTest {
    param([string]$LogFolder)
    if ([string]::IsNullOrWhiteSpace($LogFolder)) { throw '分段自测必须提供独立日志目录。' }
    $root = Join-Path $LogFolder 'composerData'
    [void](New-Item -ItemType Directory -Path $root -Force)
    try {
        # 文件名含不规律分隔：第一条规整，第二条没有连接符
        [void](New-Item -ItemType File -Path (Join-Path $root '2025-张三-王五.docx'))
        [void](New-Item -ItemType File -Path (Join-Path $root '甲X乙.txt'))

        $items = @(Get-ChildItem -LiteralPath $root -File)
        if ($items.Count -ne 2) { throw "自测失败：文件数 $($items.Count)" }

        $rows = New-Object System.Collections.ArrayList
        foreach ($item in $items) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
            $segments = Split-NameToSegments -BaseName $baseName -Separator '-' -SegmentCount 3
            if ($item.Name -eq '甲X乙.txt') {
                # 分隔符不规律：只能拆出第 1 段，后两段由用户在表格里手填
                $segments[1] = '张三'
            }
            # 模拟「段3 固定 = 李四」：勾选固定后覆盖所有行
            $segments[2] = '李四'
            [void]$rows.Add([pscustomobject]@{
                SourcePath = $item.FullName
                ParentDir  = $root
                Name       = $item.Name
                Extension  = $item.Extension
                Segments   = [string[]]$segments
            })
        }

        $plan = @(Build-ComposerPlan -Rows $rows.ToArray() -Separator '-')
        Write-Output ("PLAN={0}" -f $plan.Count)
        $plan | ForEach-Object { Write-Output ("  {0} -> {1} | {2}" -f $_.SourceName, $_.NewName, $_.Status) }

        if (@($plan | Where-Object Status -eq '有空段').Count -ne 0) { throw '自测失败：仍有空段' }
        $firstItem = @($plan | Where-Object SourceName -eq '2025-张三-王五.docx')[0]
        $secondItem = @($plan | Where-Object SourceName -eq '甲X乙.txt')[0]
        if ($firstItem.NewName -ne '2025-张三-李四.docx') { throw "自测失败：拼接结果错误 $($firstItem.NewName)" }
        if ($secondItem.NewName -ne '甲X乙-张三-李四.txt') { throw "自测失败：拼接结果错误 $($secondItem.NewName)" }

        $result = Invoke-ComposerRename -Plan $plan -LogFolder $LogFolder
        Write-Output ("RENAMED={0} ERRORS={1} OUTCOME={2}" -f $result.SuccessCount, $result.Errors.Count, $result.Outcome)
        if ($result.SuccessCount -ne 2) { throw '自测失败：改名数量不符' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '2025-张三-李四.docx'))) { throw '自测失败：缺少目标文件1' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '甲X乙-张三-李四.txt'))) { throw '自测失败：缺少目标文件2' }
        if (-not $result.LogPath -or -not (Test-Path -LiteralPath $result.LogPath)) { throw '自测失败：未生成撤销日志' }
        $logRecord = @(Import-Csv -LiteralPath $result.LogPath -Encoding UTF8)[0]
        if (-not $logRecord.OldPath -or -not $logRecord.NewPath) { throw '自测失败：日志缺少 OldPath/NewPath' }

        $undo = Invoke-UndoLatest -LogFolder $LogFolder
        Write-Output ("COMPOSER_UNDO={0} DONE={1}" -f $undo.RestoredCount, $undo.MarkedDone)
        if ($undo.RestoredCount -ne 2 -or -not $undo.MarkedDone) { throw '自测失败：分段撤销不符' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '2025-张三-王五.docx'))) { throw '自测失败：未还原文件1' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '甲X乙.txt'))) { throw '自测失败：未还原文件2' }
    } finally {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($SelfTest) {
    $officialLogs = Get-BatchRenamerLogFolder
    $fingerprintBefore = Get-BatchRenamerLogFingerprint -LogFolder $officialLogs
    $testRoot = Join-Path $env:TEMP ('BatchRenamerTests_' + [guid]::NewGuid().ToString('N'))
    $testLogs = Join-Path $testRoot 'logs'
    $failed = New-Object System.Collections.ArrayList
    try {
        [void](New-Item -ItemType Directory -Path $testLogs -Force)
        try {
            Invoke-SelfTest -LogFolder $testLogs
            Write-Output 'SUBTEST=Sequence PASS'
        } catch {
            [void]$failed.Add('Sequence: ' + $_.Exception.Message)
            Write-Output 'SUBTEST=Sequence FAIL'
            Write-Output ('ERROR: ' + $_.Exception.Message)
        }
        try {
            Invoke-ComposerSelfTest -LogFolder $testLogs
            Write-Output 'SUBTEST=Composer PASS'
        } catch {
            [void]$failed.Add('Composer: ' + $_.Exception.Message)
            Write-Output 'SUBTEST=Composer FAIL'
            Write-Output ('ERROR: ' + $_.Exception.Message)
        }
    } finally {
        if (Test-Path -LiteralPath $testRoot) {
            Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $testRoot) {
                [void]$failed.Add("未能清理测试目录：$testRoot")
                Write-Output ("WARN: 未能清理测试目录 $testRoot")
            }
        }
    }
    $fingerprintAfter = Get-BatchRenamerLogFingerprint -LogFolder $officialLogs
    if ($fingerprintBefore -ne $fingerprintAfter) {
        [void]$failed.Add('正式撤销日志在自测后发生变化')
    }
    if ($failed.Count -gt 0) {
        Write-Output 'SELFTEST=FAIL'
        foreach ($item in $failed) { Write-Output ('FAIL_ITEM: ' + $item) }
        exit 1
    }
    Write-Output 'SELFTEST=PASS'
    exit 0
}

# ---------------------------------------------------------------------------
# 界面
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$script:LogFolder = Join-Path $env:TEMP 'BatchRenamerLogs'
$script:FileItems = @()
$script:FixedChecks = @()
$script:FixedBoxes = @()
$script:SegmentCount = 3
$script:SegmentFolder = ''
# 程序内部批量刷新单元格时会触发 CellValueChanged，用该开关阻断递归刷新
$script:SuppressRefresh = $false
$script:RestoringFixed = $false
# 固定值控件上一次排布时的容器宽度，窗口尺寸变化后据此判断是否需要重排
$script:FixedPanelWidth = 0
$script:ActiveView = 'sequence'
$script:SeqPlanDirty = $true
$script:SegmentStale = $true

$font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$fontBold = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$fontSmall = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5)
$fontInput = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)
$fontNav = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$fontTitle = New-Object System.Drawing.Font('Microsoft YaHei UI', 13.5, [System.Drawing.FontStyle]::Bold)
$fontSubtitle = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$fontGridHeader = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5, [System.Drawing.FontStyle]::Bold)

function Get-Color {
    param([string]$Html)
    return [System.Drawing.ColorTranslator]::FromHtml($Html)
}

$script:Ui = @{
    Canvas   = '#F5F7FA'
    Panel    = '#FFFFFF'
    Text     = '#172033'
    Muted    = '#64748B'
    Line     = '#E2E8F0'
    Primary  = '#2563EB'
    PrimaryH = '#1D4ED8'
    NavHover = '#EFF6FF'
    NavIdle  = '#F8FAFC'
    Success  = '#059669'
    SuccessB = '#ECFDF5'
    Danger   = '#DC2626'
    DangerB  = '#FEF2F2'
    Warn     = '#D97706'
    SampleB  = '#EFF6FF'
}

# ===== 控件工厂 =====

function Initialize-ButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button,
        [ValidateSet('Default', 'Primary', 'Danger', 'Ghost')][string]$Kind = 'Default'
    )
    $Button.FlatStyle = 'Flat'
    $Button.Font = $font
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.UseVisualStyleBackColor = $false
    switch ($Kind) {
        'Primary' {
            $Button.BackColor = Get-Color $script:Ui.Primary
            $Button.ForeColor = [System.Drawing.Color]::White
            $Button.FlatAppearance.BorderSize = 1
            $Button.FlatAppearance.BorderColor = Get-Color $script:Ui.Primary
            $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.PrimaryH
        }
        'Danger' {
            $Button.BackColor = [System.Drawing.Color]::White
            $Button.ForeColor = Get-Color $script:Ui.Danger
            $Button.FlatAppearance.BorderSize = 1
            $Button.FlatAppearance.BorderColor = Get-Color '#FECACA'
            $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.DangerB
        }
        'Ghost' {
            $Button.BackColor = [System.Drawing.Color]::White
            $Button.ForeColor = Get-Color $script:Ui.Muted
            $Button.FlatAppearance.BorderSize = 0
            $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::White
            $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.NavHover
        }
        default {
            $Button.BackColor = [System.Drawing.Color]::White
            $Button.ForeColor = Get-Color $script:Ui.Text
            $Button.FlatAppearance.BorderSize = 1
            $Button.FlatAppearance.BorderColor = Get-Color $script:Ui.Line
            $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.NavIdle
        }
    }
}

function New-FieldLabel {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [System.Windows.Forms.Control]$Parent
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.AutoSize = $true
    $label.Font = $fontSmall
    $label.ForeColor = Get-Color $script:Ui.Muted
    $label.BackColor = [System.Drawing.Color]::Transparent
    $Parent.Controls.Add($label)
}

function Initialize-Input {
    param([System.Windows.Forms.Control]$Control)
    $Control.Font = $fontInput
    if ($Control -is [System.Windows.Forms.TextBox]) {
        $Control.BorderStyle = 'FixedSingle'
        $Control.BackColor = [System.Drawing.Color]::White
    }
}

function Initialize-GridStyle {
    param([System.Windows.Forms.DataGridView]$Grid)
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = 'FullRowSelect'
    $Grid.MultiSelect = $false
    $Grid.BackgroundColor = [System.Drawing.Color]::White
    $Grid.BorderStyle = 'None'
    $Grid.CellBorderStyle = 'SingleHorizontal'
    $Grid.GridColor = Get-Color $script:Ui.Line
    $Grid.RowTemplate.Height = 34
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersHeight = 36
    $Grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = Get-Color $script:Ui.NavIdle
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = Get-Color $script:Ui.Muted
    $Grid.ColumnHeadersDefaultCellStyle.Font = $fontGridHeader
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Get-Color $script:Ui.NavIdle
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = Get-Color '#F8FAFC'
    $Grid.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
    $Grid.DefaultCellStyle.SelectionBackColor = Get-Color $script:Ui.NavHover
    $Grid.DefaultCellStyle.SelectionForeColor = Get-Color $script:Ui.Text
    $Grid.ShowCellToolTips = $true
    $Grid.DefaultCellStyle.WrapMode = 'False'
    $Grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
    $Grid.DefaultCellStyle.Font = $font
}

function New-CardPanel {
    # 思路：外层 1px 深色面板 + 内层白面板模拟「卡片边框 + 圆角卡片」的观感
    param(
        [string]$Title,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [System.Windows.Forms.Control]$Parent,
        [switch]$FillHeight
    )
    $border = New-Object System.Windows.Forms.Panel
    $border.Location = New-Object System.Drawing.Point($X, $Y)
    $border.Size = New-Object System.Drawing.Size($Width, $Height)
    $border.BackColor = Get-Color $script:Ui.Line
    $border.Padding = New-Object System.Windows.Forms.Padding(1)
    $border.Anchor = if ($FillHeight) { 'Top,Bottom,Left,Right' } else { 'Top,Left,Right' }

    $inner = New-Object System.Windows.Forms.Panel
    $inner.Dock = 'Fill'
    $inner.BackColor = [System.Drawing.Color]::White
    $border.Controls.Add($inner)
    # 先挂到视图上：内层面板此时才真正拿到最终尺寸，
    # 后续带 Left,Right 锚定的子控件才能算出正确的右侧留白（否则会被拉伸到卡片外）
    $Parent.Controls.Add($border)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Location = New-Object System.Drawing.Point(18, 13)
    $titleLabel.AutoSize = $true
    $titleLabel.Font = $fontBold
    $titleLabel.ForeColor = Get-Color $script:Ui.Text
    $inner.Controls.Add($titleLabel)

    $separator = New-Object System.Windows.Forms.Panel
    $separator.Location = New-Object System.Drawing.Point(0, 42)
    $separator.Size = New-Object System.Drawing.Size(($Width - 2), 1)
    $separator.BackColor = Get-Color $script:Ui.Line
    $separator.Anchor = 'Top,Left,Right'
    $inner.Controls.Add($separator)

    $body = New-Object System.Windows.Forms.Panel
    $body.Location = New-Object System.Drawing.Point(0, 43)
    $body.Size = New-Object System.Drawing.Size(($Width - 2), ($Height - 45))
    $body.BackColor = [System.Drawing.Color]::White
    $body.Anchor = 'Top,Bottom,Left,Right'
    $inner.Controls.Add($body)

    return [pscustomobject]@{
        Border    = $border
        Inner     = $inner
        Body      = $body
        Title     = $titleLabel
        Separator = $separator
    }
}

function Set-StatTag {
    param(
        [System.Windows.Forms.Label]$Label,
        [string]$Text,
        [ValidateSet('Success', 'Muted', 'Danger')][string]$Kind = 'Muted'
    )
    $Label.Text = $Text
    switch ($Kind) {
        'Success' { $Label.BackColor = Get-Color $script:Ui.SuccessB; $Label.ForeColor = Get-Color $script:Ui.Success }
        'Danger'  { $Label.BackColor = Get-Color $script:Ui.DangerB; $Label.ForeColor = Get-Color $script:Ui.Danger }
        default   { $Label.BackColor = Get-Color '#F1F5F9'; $Label.ForeColor = Get-Color $script:Ui.Muted }
    }
    $measured = [System.Windows.Forms.TextRenderer]::MeasureText($Text, $Label.Font)
    $Label.Width = $measured.Width + 20
}

function New-StatTag {
    param([int]$X, [int]$Y, [System.Windows.Forms.Control]$Parent)
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size(80, 22)
    $label.TextAlign = 'MiddleCenter'
    $label.Font = $fontSmall
    $label.BackColor = Get-Color '#F1F5F9'
    $label.ForeColor = Get-Color $script:Ui.Muted
    $Parent.Controls.Add($label)
    return $label
}

# ===== 窗体骨架 =====
$workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$clientWidth = [math]::Min(1200, [math]::Max(900, $workArea.Width - 60))
$clientHeight = [math]::Min(880, [math]::Max(640, $workArea.Height - 70))
if ($clientWidth -gt $workArea.Width) { $clientWidth = [math]::Max(640, $workArea.Width - 20) }
if ($clientHeight -gt $workArea.Height) { $clientHeight = [math]::Max(480, $workArea.Height - 20) }

$form = New-Object System.Windows.Forms.Form
$form.Text = '批量重命名工具'
$form.ClientSize = New-Object System.Drawing.Size($clientWidth, $clientHeight)
$minWidth = [math]::Min(1080, [math]::Max(800, $workArea.Width - 16))
$minHeight = [math]::Min(760, [math]::Max(560, $workArea.Height - 16))
$form.MinimumSize = New-Object System.Drawing.Size($minWidth, $minHeight)
$form.AutoScaleMode = 'Font'
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.BackColor = Get-Color $script:Ui.Canvas

$sidebarWidth = 196
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Location = New-Object System.Drawing.Point(0, 0)
$sidebar.Size = New-Object System.Drawing.Size($sidebarWidth, $clientHeight)
$sidebar.BackColor = Get-Color $script:Ui.Canvas
$sidebar.Anchor = 'Top,Bottom,Left'
$form.Controls.Add($sidebar)

$sidebarEdge = New-Object System.Windows.Forms.Panel
$sidebarEdge.Location = New-Object System.Drawing.Point(($sidebarWidth - 1), 0)
$sidebarEdge.Size = New-Object System.Drawing.Size(1, $clientHeight)
$sidebarEdge.BackColor = Get-Color $script:Ui.Line
$sidebarEdge.Anchor = 'Top,Bottom,Right'
$sidebar.Controls.Add($sidebarEdge)

$logoBox = New-Object System.Windows.Forms.Panel
$logoBox.Location = New-Object System.Drawing.Point(16, 20)
$logoBox.Size = New-Object System.Drawing.Size(22, 22)
$logoBox.BackColor = Get-Color $script:Ui.Primary
$sidebar.Controls.Add($logoBox)

$logoMark = New-Object System.Windows.Forms.Label
$logoMark.Text = '≡'
$logoMark.Dock = 'Fill'
$logoMark.TextAlign = 'MiddleCenter'
$logoMark.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$logoMark.ForeColor = [System.Drawing.Color]::White
$logoBox.Controls.Add($logoMark)

$logoText = New-Object System.Windows.Forms.Label
$logoText.Text = '批量重命名'
$logoText.Location = New-Object System.Drawing.Point(46, 21)
$logoText.AutoSize = $true
$logoText.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$logoText.ForeColor = Get-Color $script:Ui.Text
$sidebar.Controls.Add($logoText)

$sidebarTopBorder = New-Object System.Windows.Forms.Panel
$sidebarTopBorder.Location = New-Object System.Drawing.Point(0, 62)
$sidebarTopBorder.Size = New-Object System.Drawing.Size($sidebarWidth, 1)
$sidebarTopBorder.BackColor = Get-Color $script:Ui.Line
$sidebarTopBorder.Anchor = 'Top,Left,Right'
$sidebar.Controls.Add($sidebarTopBorder)

$navLabel = New-Object System.Windows.Forms.Label
$navLabel.Text = '命名模式'
$navLabel.Location = New-Object System.Drawing.Point(16, 78)
$navLabel.AutoSize = $true
$navLabel.Font = $fontSmall
$navLabel.ForeColor = Get-Color $script:Ui.Muted
$sidebar.Controls.Add($navLabel)

function Set-NavButtonState {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$ViewName
    )
    $canvas = Get-Color $script:Ui.Canvas
    if ($script:ActiveView -eq $ViewName) {
        $Button.BackColor = Get-Color $script:Ui.Primary
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.Primary
        $Button.FlatAppearance.MouseDownBackColor = Get-Color $script:Ui.Primary
        $Button.FlatAppearance.BorderColor = Get-Color $script:Ui.Primary
    } else {
        $Button.BackColor = $canvas
        $Button.ForeColor = Get-Color $script:Ui.Muted
        $Button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.NavHover
        $Button.FlatAppearance.MouseDownBackColor = Get-Color $script:Ui.NavHover
        $Button.FlatAppearance.BorderColor = $canvas
    }
}

function New-NavButton {
    param([string]$Text, [int]$Y, [string]$ViewName)
    $canvas = Get-Color $script:Ui.Canvas
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "  $Text"
    $button.Location = New-Object System.Drawing.Point(10, $Y)
    $button.Size = New-Object System.Drawing.Size(($sidebarWidth - 22), 40)
    $button.TextAlign = 'MiddleLeft'
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderSize = 0
    $button.FlatAppearance.MouseOverBackColor = Get-Color $script:Ui.NavHover
    $button.FlatAppearance.MouseDownBackColor = Get-Color $script:Ui.NavHover
    $button.FlatAppearance.BorderColor = $canvas
    $button.BackColor = $canvas
    $button.ForeColor = Get-Color $script:Ui.Muted
    $button.Font = $fontNav
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.UseVisualStyleBackColor = $false
    $button.TabStop = $false
    $button.Tag = $ViewName
    $button.Add_Click({ Show-View ([string]$this.Tag) })
    $button.Add_MouseEnter({
        if ($script:ActiveView -ne [string]$this.Tag) { $this.BackColor = Get-Color $script:Ui.NavHover }
    })
    $button.Add_MouseLeave({
        Set-NavButtonState -Button $this -ViewName ([string]$this.Tag)
    })
    $sidebar.Controls.Add($button)
    return $button
}

$btnNavSequence = New-NavButton -Text '序号编号' -Y 102 -ViewName 'sequence'
$btnNavSegment = New-NavButton -Text '分段拼接' -Y 146 -ViewName 'segment'

$sidebarFooter = New-Object System.Windows.Forms.Label
$sidebarFooter.Text = "批量重命名工具  v1.1.0`n让文件管理更简单"
$sidebarFooter.Location = New-Object System.Drawing.Point(16, ($clientHeight - 48))
$sidebarFooter.AutoSize = $true
$sidebarFooter.Font = $fontSmall
$sidebarFooter.ForeColor = Get-Color $script:Ui.Muted
$sidebarFooter.Anchor = 'Bottom,Left'
$sidebar.Controls.Add($sidebarFooter)

# ===== 内容区 =====
$contentHost = New-Object System.Windows.Forms.Panel
$contentHost.Location = New-Object System.Drawing.Point($sidebarWidth, 0)
$contentHost.Size = New-Object System.Drawing.Size(($clientWidth - $sidebarWidth), $clientHeight)
$contentHost.BackColor = Get-Color $script:Ui.Canvas
$contentHost.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($contentHost)

$viewSequence = New-Object System.Windows.Forms.Panel
$viewSequence.Location = New-Object System.Drawing.Point(0, 0)
$viewSequence.Size = New-Object System.Drawing.Size(($clientWidth - $sidebarWidth), $clientHeight)
$viewSequence.BackColor = Get-Color $script:Ui.Canvas
$viewSequence.Anchor = 'Top,Bottom,Left,Right'
$contentHost.Controls.Add($viewSequence)

$viewSegment = New-Object System.Windows.Forms.Panel
$viewSegment.Location = New-Object System.Drawing.Point(0, 0)
$viewSegment.Size = New-Object System.Drawing.Size(($clientWidth - $sidebarWidth), $clientHeight)
$viewSegment.BackColor = Get-Color $script:Ui.Canvas
$viewSegment.Anchor = 'Top,Bottom,Left,Right'
$viewSegment.Visible = $false
$contentHost.Controls.Add($viewSegment)

$script:PathTip = New-Object System.Windows.Forms.ToolTip
$script:PathTip.AutoPopDelay = 10000
$script:PathTip.InitialDelay = 400
$script:PathTip.ReshowDelay = 100

# ----- 视图标题 -----
function Add-PageHeader {
    param([string]$Title, [string]$Subtitle, [System.Windows.Forms.Control]$Parent)
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Location = New-Object System.Drawing.Point(20, 12)
    $titleLabel.AutoSize = $true
    $titleLabel.Font = $fontTitle
    $titleLabel.ForeColor = Get-Color $script:Ui.Text
    $Parent.Controls.Add($titleLabel)

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = $Subtitle
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 40)
    $subtitleLabel.AutoSize = $true
    $subtitleLabel.Font = $fontSubtitle
    $subtitleLabel.ForeColor = Get-Color $script:Ui.Muted
    $Parent.Controls.Add($subtitleLabel)
}

function New-SplitWorkArea {
    param(
        [System.Windows.Forms.Control]$Parent,
        [int]$LeftWidth
    )
    $headerH = 66
    $work = New-Object System.Windows.Forms.Panel
    $work.Location = New-Object System.Drawing.Point(0, $headerH)
    $work.Size = New-Object System.Drawing.Size($Parent.ClientSize.Width, ($Parent.ClientSize.Height - $headerH))
    $work.BackColor = Get-Color $script:Ui.Canvas
    $work.Anchor = 'Top,Bottom,Left,Right'
    $Parent.Controls.Add($work)

    $left = New-Object System.Windows.Forms.Panel
    $left.Dock = 'Left'
    $left.Width = $LeftWidth
    $left.BackColor = Get-Color $script:Ui.Canvas
    $left.AutoScroll = $true

    $right = New-Object System.Windows.Forms.Panel
    $right.Dock = 'Fill'
    $right.BackColor = Get-Color $script:Ui.Canvas

    $work.Controls.Add($right)
    $work.Controls.Add($left)
    return [pscustomobject]@{ Work = $work; Left = $left; Right = $right }
}

Add-PageHeader -Title '序号编号' -Subtitle '为文件或文件夹按规则批量添加序号' -Parent $viewSequence
Add-PageHeader -Title '分段拼接' -Subtitle '按规则将文件名分段，重新组合' -Parent $viewSegment

$viewContentWidth = $clientWidth - $sidebarWidth
$leftColWidth = 508
$seqSplit = New-SplitWorkArea -Parent $viewSequence -LeftWidth $leftColWidth
$segSplit = New-SplitWorkArea -Parent $viewSegment -LeftWidth $leftColWidth
$leftCardWidth = $leftColWidth - 32
$cardWidth = $leftCardWidth

# ---------------------------------------------------------------------------
# 视图1：序号编号
# ---------------------------------------------------------------------------

# 卡片1：选择目录
$cardFolderSeq = New-CardPanel -Title '1. 选择目录' -X 16 -Y 8 -Width $leftCardWidth -Height 96 -Parent $seqSplit.Left
$seqBodyWidth = $leftCardWidth - 2

$txtSeqFolder = New-Object System.Windows.Forms.TextBox
$txtSeqFolder.Location = New-Object System.Drawing.Point(14, 8)
$txtSeqFolder.Size = New-Object System.Drawing.Size(($seqBodyWidth - 122), 26)
$txtSeqFolder.Anchor = 'Top,Left,Right'
Initialize-Input -Control $txtSeqFolder
$cardFolderSeq.Body.Controls.Add($txtSeqFolder)

$btnSeqBrowse = New-Object System.Windows.Forms.Button
$btnSeqBrowse.Text = '浏览...'
$btnSeqBrowse.Location = New-Object System.Drawing.Point(($seqBodyWidth - 96), 7)
$btnSeqBrowse.Size = New-Object System.Drawing.Size(80, 28)
$btnSeqBrowse.Anchor = 'Top,Right'
$cardFolderSeq.Body.Controls.Add($btnSeqBrowse)
Initialize-ButtonStyle -Button $btnSeqBrowse

# 卡片2：命名规则
$cardRule = New-CardPanel -Title '2. 命名规则' -X 16 -Y 116 -Width $leftCardWidth -Height 318 -Parent $seqSplit.Left
$ruleBody = $cardRule.Body

New-FieldLabel -Text '模式' -X 14 -Y 2 -Parent $ruleBody
$cmbNameMode = New-Object System.Windows.Forms.ComboBox
$cmbNameMode.DropDownStyle = 'DropDownList'
$cmbNameMode.Location = New-Object System.Drawing.Point(14, 22)
$cmbNameMode.Size = New-Object System.Drawing.Size(($seqBodyWidth - 28), 26)
$cmbNameMode.Anchor = 'Top,Left,Right'
[void]$cmbNameMode.Items.Add('序号命名（前缀+序号+后缀）')
[void]$cmbNameMode.Items.Add('仅加前缀（保留原名）')
[void]$cmbNameMode.Items.Add('仅加后缀（保留原名）')
[void]$cmbNameMode.Items.Add('前缀+序号+原名')
$cmbNameMode.SelectedIndex = 0
$cmbNameMode.Font = $fontInput
$ruleBody.Controls.Add($cmbNameMode)

New-FieldLabel -Text '前缀' -X 14 -Y 54 -Parent $ruleBody
$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Text = '文件_'
$txtPrefix.Location = New-Object System.Drawing.Point(14, 74)
$txtPrefix.Size = New-Object System.Drawing.Size(140, 26)
Initialize-Input -Control $txtPrefix
$ruleBody.Controls.Add($txtPrefix)

New-FieldLabel -Text '后缀' -X 168 -Y 54 -Parent $ruleBody
$txtSuffix = New-Object System.Windows.Forms.TextBox
$txtSuffix.Location = New-Object System.Drawing.Point(168, 74)
$txtSuffix.Size = New-Object System.Drawing.Size(140, 26)
Initialize-Input -Control $txtSuffix
$ruleBody.Controls.Add($txtSuffix)

New-FieldLabel -Text '起始序号' -X 322 -Y 54 -Parent $ruleBody
$numStart = New-Object System.Windows.Forms.NumericUpDown
$numStart.Location = New-Object System.Drawing.Point(322, 74)
$numStart.Size = New-Object System.Drawing.Size(70, 26)
$numStart.Minimum = 0
$numStart.Maximum = 999999
$numStart.Value = 1
$numStart.Font = $fontInput
$ruleBody.Controls.Add($numStart)

New-FieldLabel -Text '位数' -X 402 -Y 54 -Parent $ruleBody
$numDigits = New-Object System.Windows.Forms.NumericUpDown
$numDigits.Location = New-Object System.Drawing.Point(402, 74)
$numDigits.Size = New-Object System.Drawing.Size(56, 26)
$numDigits.Minimum = 1
$numDigits.Maximum = 10
$numDigits.Value = 3
$numDigits.Font = $fontInput
$ruleBody.Controls.Add($numDigits)

New-FieldLabel -Text '排序方式' -X 14 -Y 106 -Parent $ruleBody
$cmbSort = New-Object System.Windows.Forms.ComboBox
$cmbSort.DropDownStyle = 'DropDownList'
$cmbSort.Location = New-Object System.Drawing.Point(14, 126)
$cmbSort.Size = New-Object System.Drawing.Size(210, 26)
[void]$cmbSort.Items.Add('名称（自然序）')
[void]$cmbSort.Items.Add('修改时间')
[void]$cmbSort.Items.Add('创建时间')
[void]$cmbSort.Items.Add('大小')
[void]$cmbSort.Items.Add('随机')
$cmbSort.SelectedIndex = 0
$cmbSort.Font = $fontInput
$ruleBody.Controls.Add($cmbSort)

# 查找替换行：标签宽度固定，说明文字单独放在最右侧并裁切，避免长文本压住复选框
$lblReplaceTitle = New-Object System.Windows.Forms.Label
$lblReplaceTitle.Text = '查找替换'
$lblReplaceTitle.Location = New-Object System.Drawing.Point(14, 160)
$lblReplaceTitle.AutoSize = $false
$lblReplaceTitle.Size = New-Object System.Drawing.Size(56, 20)
$lblReplaceTitle.Font = $fontSmall
$lblReplaceTitle.ForeColor = Get-Color $script:Ui.Muted
$ruleBody.Controls.Add($lblReplaceTitle)

$chkReplace = New-Object System.Windows.Forms.CheckBox
$chkReplace.Text = '启用'
$chkReplace.Location = New-Object System.Drawing.Point(74, 157)
$chkReplace.Size = New-Object System.Drawing.Size(56, 24)
$chkReplace.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($chkReplace)

$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = '查找'
$lblFind.Location = New-Object System.Drawing.Point(138, 160)
$lblFind.AutoSize = $true
$lblFind.ForeColor = Get-Color $script:Ui.Muted
$lblFind.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(172, 156)
$txtFind.Size = New-Object System.Drawing.Size(100, 26)
Initialize-Input -Control $txtFind
$ruleBody.Controls.Add($txtFind)

$lblArrow = New-Object System.Windows.Forms.Label
$lblArrow.Text = '→'
$lblArrow.Location = New-Object System.Drawing.Point(276, 160)
$lblArrow.AutoSize = $true
$lblArrow.ForeColor = Get-Color $script:Ui.Muted
$lblArrow.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($lblArrow)

$lblReplaceTo = New-Object System.Windows.Forms.Label
$lblReplaceTo.Text = '替换为'
$lblReplaceTo.Location = New-Object System.Drawing.Point(296, 160)
$lblReplaceTo.AutoSize = $true
$lblReplaceTo.ForeColor = Get-Color $script:Ui.Muted
$lblReplaceTo.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($lblReplaceTo)

$txtReplaceNew = New-Object System.Windows.Forms.TextBox
$txtReplaceNew.Location = New-Object System.Drawing.Point(350, 156)
$txtReplaceNew.Size = New-Object System.Drawing.Size(110, 26)
Initialize-Input -Control $txtReplaceNew
$ruleBody.Controls.Add($txtReplaceNew)

$chkCase = New-Object System.Windows.Forms.CheckBox
$chkCase.Text = '区分大小写'
$chkCase.Location = New-Object System.Drawing.Point(14, 188)
$chkCase.Size = New-Object System.Drawing.Size(100, 24)
$chkCase.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($chkCase)

$lblReplaceHint = New-Object System.Windows.Forms.Label
$lblReplaceHint.Text = '作用于原名称，不含扩展名'
$lblReplaceHint.Location = New-Object System.Drawing.Point(120, 192)
$lblReplaceHint.Size = New-Object System.Drawing.Size(($seqBodyWidth - 134), 18)
$lblReplaceHint.Anchor = 'Top,Left,Right'
$lblReplaceHint.AutoSize = $false
$lblReplaceHint.AutoEllipsis = $true
$lblReplaceHint.Font = $fontSmall
$lblReplaceHint.ForeColor = Get-Color $script:Ui.Muted
$lblReplaceHint.BackColor = [System.Drawing.Color]::Transparent
$ruleBody.Controls.Add($lblReplaceHint)

$lblSample = New-Object System.Windows.Forms.Label
$lblSample.Location = New-Object System.Drawing.Point(14, 222)
$lblSample.Size = New-Object System.Drawing.Size(($seqBodyWidth - 28), 34)
$lblSample.Anchor = 'Top,Left,Right'
$lblSample.BackColor = Get-Color $script:Ui.SampleB
$lblSample.ForeColor = Get-Color $script:Ui.Primary
$lblSample.Font = $font
$lblSample.TextAlign = 'MiddleLeft'
$lblSample.Padding = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
$ruleBody.Controls.Add($lblSample)

# 卡片3：筛选设置
$cardScope = New-CardPanel -Title '3. 筛选设置' -X 16 -Y 446 -Width $leftCardWidth -Height 164 -Parent $seqSplit.Left
$scopeBody = $cardScope.Body

$rbFiles = New-Object System.Windows.Forms.RadioButton
$rbFiles.Text = '文件'
$rbFiles.Location = New-Object System.Drawing.Point(14, 10)
$rbFiles.Size = New-Object System.Drawing.Size(64, 24)
$rbFiles.Checked = $true
$rbFiles.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($rbFiles)

$rbFolders = New-Object System.Windows.Forms.RadioButton
$rbFolders.Text = '文件夹'
$rbFolders.Location = New-Object System.Drawing.Point(86, 10)
$rbFolders.Size = New-Object System.Drawing.Size(82, 24)
$rbFolders.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($rbFolders)

$rbBoth = New-Object System.Windows.Forms.RadioButton
$rbBoth.Text = '两者'
$rbBoth.Location = New-Object System.Drawing.Point(174, 10)
$rbBoth.Size = New-Object System.Drawing.Size(64, 24)
$rbBoth.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($rbBoth)

$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text = '含子目录'
$chkRecurse.Location = New-Object System.Drawing.Point(250, 10)
$chkRecurse.Size = New-Object System.Drawing.Size(90, 24)
$chkRecurse.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($chkRecurse)

$lblRecurseHint = New-Object System.Windows.Forms.Label
$lblRecurseHint.Text = '递归开启后只处理文件，文件夹选项会禁用'
$lblRecurseHint.Location = New-Object System.Drawing.Point(14, 36)
$lblRecurseHint.AutoSize = $true
$lblRecurseHint.Font = $fontSmall
$lblRecurseHint.ForeColor = Get-Color $script:Ui.Muted
$lblRecurseHint.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($lblRecurseHint)

New-FieldLabel -Text '扩展名' -X 14 -Y 62 -Parent $scopeBody
$txtExtFilter = New-Object System.Windows.Forms.TextBox
$txtExtFilter.Location = New-Object System.Drawing.Point(70, 58)
$txtExtFilter.Size = New-Object System.Drawing.Size(150, 26)
Initialize-Input -Control $txtExtFilter
$scopeBody.Controls.Add($txtExtFilter)

$lblExtHint = New-Object System.Windows.Forms.Label
$lblExtHint.Text = '如 jpg,png（留空=全部）'
$lblExtHint.Location = New-Object System.Drawing.Point(14, 88)
$lblExtHint.AutoSize = $true
$lblExtHint.Font = $fontSmall
$lblExtHint.ForeColor = Get-Color $script:Ui.Muted
$lblExtHint.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($lblExtHint)

$rbInclude = New-Object System.Windows.Forms.RadioButton
$rbInclude.Text = '仅这些'
$rbInclude.Location = New-Object System.Drawing.Point(232, 60)
$rbInclude.Size = New-Object System.Drawing.Size(76, 24)
$rbInclude.Checked = $true
$rbInclude.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($rbInclude)

$rbExclude = New-Object System.Windows.Forms.RadioButton
$rbExclude.Text = '排除'
$rbExclude.Location = New-Object System.Drawing.Point(314, 60)
$rbExclude.Size = New-Object System.Drawing.Size(64, 24)
$rbExclude.BackColor = [System.Drawing.Color]::Transparent
$scopeBody.Controls.Add($rbExclude)

# 卡片4：预览结果（右侧主栏）
$previewCardWidth = [math]::Max(280, $viewContentWidth - $leftColWidth - 24)
$previewCardHeight = [math]::Max(280, $clientHeight - 66 - 16)
$cardPreview = New-CardPanel -Title '预览结果' -X 8 -Y 0 -Width $previewCardWidth -Height $previewCardHeight -Parent $seqSplit.Right -FillHeight
$previewBody = $cardPreview.Body
$previewBodyWidth = $previewCardWidth - 2
$previewBodyHeight = $previewCardHeight - 45

$statsBar = New-Object System.Windows.Forms.Panel
$statsBar.Location = New-Object System.Drawing.Point(0, 0)
$statsBar.Size = New-Object System.Drawing.Size($previewBodyWidth, 38)
$statsBar.BackColor = Get-Color $script:Ui.NavIdle
$statsBar.Anchor = 'Top,Left,Right'
$previewBody.Controls.Add($statsBar)

$statsBarLine = New-Object System.Windows.Forms.Panel
$statsBarLine.Location = New-Object System.Drawing.Point(0, 37)
$statsBarLine.Size = New-Object System.Drawing.Size($previewBodyWidth, 1)
$statsBarLine.BackColor = Get-Color $script:Ui.Line
$statsBarLine.Anchor = 'Bottom,Left,Right'
$statsBar.Controls.Add($statsBarLine)

$lblSeqCount = New-Object System.Windows.Forms.Label
$lblSeqCount.Text = '尚未预览'
$lblSeqCount.Location = New-Object System.Drawing.Point(16, 9)
$lblSeqCount.AutoSize = $true
$lblSeqCount.Font = $fontSmall
$lblSeqCount.ForeColor = Get-Color $script:Ui.Muted
$statsBar.Controls.Add($lblSeqCount)

$lblSeqHint = New-Object System.Windows.Forms.Label
$lblSeqHint.Text = '请先预览'
$lblSeqHint.AutoSize = $true
$lblSeqHint.Font = $fontSmall
$lblSeqHint.ForeColor = Get-Color $script:Ui.Primary
$lblSeqHint.Anchor = 'Top,Right'
$lblSeqHint.Location = New-Object System.Drawing.Point(($previewBodyWidth - 90), 9)
$statsBar.Controls.Add($lblSeqHint)

$lblSeqExec = New-StatTag -X 120 -Y 8 -Parent $statsBar
$lblSeqUnchanged = New-StatTag -X 220 -Y 8 -Parent $statsBar
$lblSeqBlocked = New-StatTag -X 330 -Y 8 -Parent $statsBar
Set-StatTag -Label $lblSeqExec -Text '可执行 0' -Kind 'Muted'
Set-StatTag -Label $lblSeqUnchanged -Text '未变化 0' -Kind 'Muted'
Set-StatTag -Label $lblSeqBlocked -Text '有问题 0' -Kind 'Muted'

$actionBar = New-Object System.Windows.Forms.Panel
$actionBar.Location = New-Object System.Drawing.Point(0, ($previewBodyHeight - 52))
$actionBar.Size = New-Object System.Drawing.Size($previewBodyWidth, 52)
$actionBar.BackColor = Get-Color $script:Ui.NavIdle
$actionBar.Anchor = 'Bottom,Left,Right'
$previewBody.Controls.Add($actionBar)

$actionBarLine = New-Object System.Windows.Forms.Panel
$actionBarLine.Location = New-Object System.Drawing.Point(0, 0)
$actionBarLine.Size = New-Object System.Drawing.Size($previewBodyWidth, 1)
$actionBarLine.BackColor = Get-Color $script:Ui.Line
$actionBarLine.Anchor = 'Top,Left,Right'
$actionBar.Controls.Add($actionBarLine)

$btnSeqExecute = New-Object System.Windows.Forms.Button
$btnSeqExecute.Text = '执行重命名'
$btnSeqExecute.Location = New-Object System.Drawing.Point(($previewBodyWidth - 134), 11)
$btnSeqExecute.Size = New-Object System.Drawing.Size(120, 30)
$btnSeqExecute.Anchor = 'Bottom,Right'
$actionBar.Controls.Add($btnSeqExecute)
Initialize-ButtonStyle -Button $btnSeqExecute -Kind 'Primary'
$btnSeqExecute.Enabled = $false

$btnSeqPreview = New-Object System.Windows.Forms.Button
$btnSeqPreview.Text = '预览'
$btnSeqPreview.Location = New-Object System.Drawing.Point(($previewBodyWidth - 234), 11)
$btnSeqPreview.Size = New-Object System.Drawing.Size(90, 30)
$btnSeqPreview.Anchor = 'Bottom,Right'
$actionBar.Controls.Add($btnSeqPreview)
Initialize-ButtonStyle -Button $btnSeqPreview

$btnSeqUndo = New-Object System.Windows.Forms.Button
$btnSeqUndo.Text = '撤销'
$btnSeqUndo.Location = New-Object System.Drawing.Point(($previewBodyWidth - 364), 11)
$btnSeqUndo.Size = New-Object System.Drawing.Size(120, 30)
$btnSeqUndo.Anchor = 'Bottom,Right'
$actionBar.Controls.Add($btnSeqUndo)
Initialize-ButtonStyle -Button $btnSeqUndo

$gridSeq = New-Object System.Windows.Forms.DataGridView
$gridSeq.Location = New-Object System.Drawing.Point(0, 38)
$gridSeq.Size = New-Object System.Drawing.Size($previewBodyWidth, ($previewBodyHeight - 38 - 52))
$gridSeq.Anchor = 'Top,Bottom,Left,Right'
$gridSeq.ReadOnly = $true
$gridSeq.AutoSizeColumnsMode = 'Fill'
Initialize-GridStyle -Grid $gridSeq
$previewBody.Controls.Add($gridSeq)

# ---------------------------------------------------------------------------
# 视图2：分段拼接
# ---------------------------------------------------------------------------

# 卡片1：选择目录
$cardFolderSeg = New-CardPanel -Title '1. 选择目录' -X 16 -Y 8 -Width $leftCardWidth -Height 96 -Parent $segSplit.Left
$segBodyWidth = $leftCardWidth - 2

$txtSegFolder = New-Object System.Windows.Forms.TextBox
$txtSegFolder.Location = New-Object System.Drawing.Point(14, 8)
$txtSegFolder.Size = New-Object System.Drawing.Size(($segBodyWidth - 122), 26)
$txtSegFolder.Anchor = 'Top,Left,Right'
Initialize-Input -Control $txtSegFolder
$cardFolderSeg.Body.Controls.Add($txtSegFolder)

$btnSegBrowse = New-Object System.Windows.Forms.Button
$btnSegBrowse.Text = '浏览...'
$btnSegBrowse.Location = New-Object System.Drawing.Point(($segBodyWidth - 96), 7)
$btnSegBrowse.Size = New-Object System.Drawing.Size(80, 28)
$btnSegBrowse.Anchor = 'Top,Right'
$cardFolderSeg.Body.Controls.Add($btnSegBrowse)
Initialize-ButtonStyle -Button $btnSegBrowse
$script:PathTip.SetToolTip($txtSeqFolder, '')
$txtSeqFolder.Add_TextChanged({ $script:PathTip.SetToolTip($txtSeqFolder, $txtSeqFolder.Text) })
$txtSegFolder.Add_TextChanged({ $script:PathTip.SetToolTip($txtSegFolder, $txtSegFolder.Text) })

# 启动时带了 -Folder 就回填到两个视图，避免用户重复选择路径
if (-not [string]::IsNullOrWhiteSpace($Folder)) {
    $txtSeqFolder.Text = $Folder
    $txtSegFolder.Text = $Folder
}

# 卡片2：分段规则 + 命名规则
$cardSegmentRule = New-CardPanel -Title '2. 分段规则' -X 16 -Y 116 -Width $leftCardWidth -Height 118 -Parent $segSplit.Left
$segmentRuleBody = $cardSegmentRule.Body

# 标签用短文案，避免「文件名由几段组成」过长压住右侧「连接符」
New-FieldLabel -Text '段数' -X 14 -Y 2 -Parent $segmentRuleBody
$numCount = New-Object System.Windows.Forms.NumericUpDown
$numCount.Location = New-Object System.Drawing.Point(14, 22)
$numCount.Size = New-Object System.Drawing.Size(60, 26)
$numCount.Minimum = 1
$numCount.Maximum = 6
$numCount.Value = 3
$numCount.Font = $fontInput
$segmentRuleBody.Controls.Add($numCount)

New-FieldLabel -Text '连接符' -X 90 -Y 2 -Parent $segmentRuleBody
$txtSeparator = New-Object System.Windows.Forms.TextBox
$txtSeparator.Text = '-'
$txtSeparator.Location = New-Object System.Drawing.Point(90, 22)
$txtSeparator.Size = New-Object System.Drawing.Size(60, 26)
Initialize-Input -Control $txtSeparator
$segmentRuleBody.Controls.Add($txtSeparator)

$btnSplit = New-Object System.Windows.Forms.Button
$btnSplit.Text = '按连接符拆分'
$btnSplit.Location = New-Object System.Drawing.Point(166, 21)
$btnSplit.Size = New-Object System.Drawing.Size(110, 28)
$segmentRuleBody.Controls.Add($btnSplit)
Initialize-ButtonStyle -Button $btnSplit

$btnClearSegments = New-Object System.Windows.Forms.Button
$btnClearSegments.Text = '清空段值'
$btnClearSegments.Location = New-Object System.Drawing.Point(286, 21)
$btnClearSegments.Size = New-Object System.Drawing.Size(90, 28)
$segmentRuleBody.Controls.Add($btnClearSegments)
Initialize-ButtonStyle -Button $btnClearSegments -Kind 'Ghost'

$lblSegmentHint = New-Object System.Windows.Forms.Label
$lblSegmentHint.Text = '勾选「固定」可自动填入所有行，单元格仍可单独手动修改。'
$lblSegmentHint.Location = New-Object System.Drawing.Point(14, 52)
$lblSegmentHint.AutoSize = $true
$lblSegmentHint.Font = $fontSmall
$lblSegmentHint.ForeColor = Get-Color $script:Ui.Muted
$segmentRuleBody.Controls.Add($lblSegmentHint)

$cardFixed = New-CardPanel -Title '3. 固定填充' -X 16 -Y 246 -Width $leftCardWidth -Height 118 -Parent $segSplit.Left
$pnlFixedValues = New-Object System.Windows.Forms.Panel
$pnlFixedValues.Location = New-Object System.Drawing.Point(0, 8)
$pnlFixedValues.Size = New-Object System.Drawing.Size($segBodyWidth, 56)
$pnlFixedValues.BackColor = [System.Drawing.Color]::White
$pnlFixedValues.Anchor = 'Top,Left,Right'
$cardFixed.Body.Controls.Add($pnlFixedValues)

# 卡片3：预览结果（右侧主栏，内含分段编辑）
$segmentCardWidth = [math]::Max(360, $viewContentWidth - $leftColWidth - 24)
$segmentCardHeight = [math]::Max(280, $clientHeight - 66 - 16)
$cardSegmentGrid = New-CardPanel -Title '预览结果' -X 8 -Y 0 -Width $segmentCardWidth -Height $segmentCardHeight -Parent $segSplit.Right -FillHeight
$segmentGridBody = $cardSegmentGrid.Body
$segmentBodyWidth = $segmentCardWidth - 2
$segmentBodyHeight = $segmentCardHeight - 45

$segStatsBar = New-Object System.Windows.Forms.Panel
$segStatsBar.Location = New-Object System.Drawing.Point(0, 0)
$segStatsBar.Size = New-Object System.Drawing.Size($segmentBodyWidth, 38)
$segStatsBar.BackColor = Get-Color $script:Ui.NavIdle
$segStatsBar.Anchor = 'Top,Left,Right'
$segmentGridBody.Controls.Add($segStatsBar)

$segStatsBarLine = New-Object System.Windows.Forms.Panel
$segStatsBarLine.Location = New-Object System.Drawing.Point(0, 37)
$segStatsBarLine.Size = New-Object System.Drawing.Size($segmentBodyWidth, 1)
$segStatsBarLine.BackColor = Get-Color $script:Ui.Line
$segStatsBarLine.Anchor = 'Bottom,Left,Right'
$segStatsBar.Controls.Add($segStatsBarLine)

$lblSegCount = New-Object System.Windows.Forms.Label
$lblSegCount.Text = '共 0 个文件'
$lblSegCount.Location = New-Object System.Drawing.Point(16, 9)
$lblSegCount.AutoSize = $true
$lblSegCount.Font = $fontSmall
$lblSegCount.ForeColor = Get-Color $script:Ui.Muted
$segStatsBar.Controls.Add($lblSegCount)

$lblSegExec = New-StatTag -X 120 -Y 8 -Parent $segStatsBar
$lblSegUnchanged = New-StatTag -X 220 -Y 8 -Parent $segStatsBar
$lblSegBlocked = New-StatTag -X 330 -Y 8 -Parent $segStatsBar
Set-StatTag -Label $lblSegExec -Text '可执行 0' -Kind 'Muted'
Set-StatTag -Label $lblSegUnchanged -Text '未变化 0' -Kind 'Muted'
Set-StatTag -Label $lblSegBlocked -Text '有问题 0' -Kind 'Muted'

$segActionBar = New-Object System.Windows.Forms.Panel
$segActionBar.Location = New-Object System.Drawing.Point(0, ($segmentBodyHeight - 52))
$segActionBar.Size = New-Object System.Drawing.Size($segmentBodyWidth, 52)
$segActionBar.BackColor = Get-Color $script:Ui.NavIdle
$segActionBar.Anchor = 'Bottom,Left,Right'
$segmentGridBody.Controls.Add($segActionBar)

$segActionBarLine = New-Object System.Windows.Forms.Panel
$segActionBarLine.Location = New-Object System.Drawing.Point(0, 0)
$segActionBarLine.Size = New-Object System.Drawing.Size($segmentBodyWidth, 1)
$segActionBarLine.BackColor = Get-Color $script:Ui.Line
$segActionBarLine.Anchor = 'Top,Left,Right'
$segActionBar.Controls.Add($segActionBarLine)

$btnSegRun = New-Object System.Windows.Forms.Button
$btnSegRun.Text = '执行改名'
$btnSegRun.Location = New-Object System.Drawing.Point(($segmentBodyWidth - 134), 11)
$btnSegRun.Size = New-Object System.Drawing.Size(120, 30)
$btnSegRun.Anchor = 'Bottom,Right'
$segActionBar.Controls.Add($btnSegRun)
Initialize-ButtonStyle -Button $btnSegRun -Kind 'Primary'

$btnSegRefresh = New-Object System.Windows.Forms.Button
$btnSegRefresh.Text = '重新读取'
$btnSegRefresh.Location = New-Object System.Drawing.Point(($segmentBodyWidth - 244), 11)
$btnSegRefresh.Size = New-Object System.Drawing.Size(100, 30)
$btnSegRefresh.Anchor = 'Bottom,Right'
$segActionBar.Controls.Add($btnSegRefresh)
Initialize-ButtonStyle -Button $btnSegRefresh

$gridSeg = New-Object System.Windows.Forms.DataGridView
$gridSeg.Location = New-Object System.Drawing.Point(0, 38)
$gridSeg.Size = New-Object System.Drawing.Size($segmentBodyWidth, ($segmentBodyHeight - 38 - 52))
$gridSeg.Anchor = 'Top,Bottom,Left,Right'
$gridSeg.EditMode = 'EditOnEnter'
$gridSeg.AutoSizeColumnsMode = 'None'
Initialize-GridStyle -Grid $gridSeg
$segmentGridBody.Controls.Add($gridSeg)

# ---------------------------------------------------------------------------
# 交互逻辑
# ---------------------------------------------------------------------------

function Show-View {
    param([string]$Name)
    $script:ActiveView = $Name
    $viewSequence.Visible = ($Name -eq 'sequence')
    $viewSegment.Visible = ($Name -eq 'segment')

    Set-NavButtonState -Button $btnNavSequence -ViewName 'sequence'
    Set-NavButtonState -Button $btnNavSegment -ViewName 'segment'

    # 切换过去时若分段视图还没有文件夹，沿用主视图的路径，避免重复选择
    if ($Name -eq 'segment') {
        $seqFolder = $txtSeqFolder.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($txtSegFolder.Text) -and $seqFolder) {
            $txtSegFolder.Text = $seqFolder
        }
        $segFolder = $txtSegFolder.Text.Trim()
        $needLoad = $false
        if ($segFolder) {
            if ($gridSeg.Rows.Count -eq 0) { $needLoad = $true }
            if ($script:SegmentStale) { $needLoad = $true }
        }
        if ($needLoad) {
            try {
                Load-SegmentFiles -FolderPath $segFolder
                Rebuild-SegmentGrid
                [void](Refresh-SegmentStatus)
                $script:SegmentStale = $false
            } catch {
                $lblSegCount.Text = $_.Exception.Message
            }
        }
    }
}

function Update-SequenceSample {
    $digits = [int]$numDigits.Value
    $first = ([int]$numStart.Value).ToString().PadLeft($digits, '0')
    $second = ([int]$numStart.Value + 1).ToString().PadLeft($digits, '0')
    $prefixText = $txtPrefix.Text
    $suffixText = $txtSuffix.Text
    switch ($cmbNameMode.SelectedIndex) {
        1 {
            $firstText = $prefixText + '原名称.jpg'
            $secondText = '（不编号，按原名称排序）'
        }
        2 {
            $firstText = '原名称' + $suffixText + '.jpg'
            $secondText = '（不编号，按原名称排序）'
        }
        3 {
            $firstText = $prefixText + $first + $suffixText + '原名称.jpg'
            $secondText = $prefixText + $second + $suffixText + '原名称.jpg'
        }
        default {
            $firstText = $prefixText + $first + $suffixText + '.jpg'
            $secondText = $prefixText + $second + $suffixText + '.jpg'
        }
    }
    $text = "示例：$firstText    第 2 个：$secondText"
    if ($chkReplace.Checked -and $cmbNameMode.SelectedIndex -ne 0) {
        $text += '    查找替换作用于原名称'
    }
    $lblSample.Text = $text
}

function Update-ReplaceAvailability {
    # 思路：序号命名模式完全丢弃原名称，查找替换必然无效，直接禁用并说明，避免用户误判「没反应」
    $usesOriginalName = ($cmbNameMode.SelectedIndex -ne 0)
    foreach ($control in @($chkReplace, $lblFind, $txtFind, $lblArrow, $lblReplaceTo, $txtReplaceNew, $chkCase)) {
        $control.Enabled = $usesOriginalName
    }
    if ($usesOriginalName) {
        $lblReplaceHint.Text = '作用于原名称，不含扩展名'
        $lblReplaceHint.ForeColor = Get-Color $script:Ui.Muted
    } else {
        $lblReplaceHint.Text = '当前模式不使用原名称，不可用'
        $lblReplaceHint.ForeColor = Get-Color '#94A3B8'
    }
}

function Set-SeqHint {
    param([string]$Text, [ValidateSet('Info', 'Warn', 'Ok')][string]$Kind = 'Info')
    if (-not $lblSeqHint) { return }
    $lblSeqHint.Text = $Text
    switch ($Kind) {
        'Warn' { $lblSeqHint.ForeColor = Get-Color $script:Ui.Warn }
        'Ok'   { $lblSeqHint.ForeColor = Get-Color $script:Ui.Success }
        default { $lblSeqHint.ForeColor = Get-Color $script:Ui.Primary }
    }
    if ($statsBar) {
        $lblSeqHint.Left = [math]::Max(220, $statsBar.Width - $lblSeqHint.PreferredWidth - 16)
    }
}

function Get-SelectedMode {
    if ($rbFolders.Checked) { return 'Folders' }
    if ($rbBoth.Checked) { return 'Both' }
    return 'Files'
}

function Update-UndoButton {
    $count = Get-UndoCount
    $btnSeqUndo.Enabled = ($count -gt 0)
    $btnSeqUndo.Text = if ($count -gt 0) { "撤销（剩 $count 步）" } else { '撤销' }
}

function Set-SequencePlanDirty {
    $script:SeqPlanDirty = $true
    if ($btnSeqExecute) {
        $btnSeqExecute.Enabled = $false
        if ($script:Plan.Count -gt 0) {
            $lblSeqCount.Text = '参数已变化，请重新预览。'
            Set-SeqHint -Text '规则已改' -Kind 'Warn'
        } elseif ($lblSeqCount) {
            Set-SeqHint -Text '请先预览' -Kind 'Info'
        }
    }
}

function Update-SequenceExecuteState {
    $blocked = @($script:Plan | Where-Object { $script:BlockedStatuses -contains $_.Status }).Count
    $executable = @($script:Plan | Where-Object Status -eq '可执行').Count
    $canRun = (-not $script:SeqPlanDirty) -and ($script:Plan.Count -gt 0) -and ($executable -gt 0) -and ($blocked -eq 0)
    $btnSeqExecute.Enabled = $canRun
}

function Notify-DirectoryChanged {
    $script:Plan = @()
    $script:SeqPlanDirty = $true
    $script:SegmentStale = $true
    if ($gridSeq) { $gridSeq.DataSource = $null }
    if ($lblSeqCount) { $lblSeqCount.Text = '尚未预览' }
    Set-SeqHint -Text '请先预览' -Kind 'Info'
    Update-SequenceExecuteState
    Update-UndoButton
}

function Update-SequenceStats {
    param([int]$Total, [int]$Executable, [int]$Unchanged, [int]$Blocked)
    $lblSeqCount.Text = "共 $Total 项"
    Set-StatTag -Label $lblSeqExec -Text "可执行 $Executable" -Kind $(if ($Executable -gt 0) { 'Success' } else { 'Muted' })
    Set-StatTag -Label $lblSeqUnchanged -Text "未变化 $Unchanged" -Kind 'Muted'
    Set-StatTag -Label $lblSeqBlocked -Text "有问题 $Blocked" -Kind $(if ($Blocked -gt 0) { 'Danger' } else { 'Muted' })
    $lblSeqExec.Location = New-Object System.Drawing.Point((120), 8)
    $lblSeqUnchanged.Location = New-Object System.Drawing.Point((120 + $lblSeqExec.Width + 10), 8)
    $lblSeqBlocked.Location = New-Object System.Drawing.Point((120 + $lblSeqExec.Width + 10 + $lblSeqUnchanged.Width + 10), 8)
}

function Update-SequencePreview {
    param([string]$FolderPath)
    # 必须先判空：Test-Path -LiteralPath 收到空字符串时参数绑定就抛错
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        throw '请先选择目标文件夹（点右侧「浏览...」）。'
    }
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        throw '目标文件夹不存在，请重新选择。'
    }
    $sortMap = @('Name', 'LastWriteTime', 'CreationTime', 'Size', 'Random')
    $nameModeMap = @('Numbered', 'PrefixOnly', 'SuffixOnly', 'PrefixNumberOriginal')
    $sortBy = $sortMap[$cmbSort.SelectedIndex]
    $nameMode = $nameModeMap[$cmbNameMode.SelectedIndex]
    $replaceFrom = if ($chkReplace.Checked -and $nameMode -ne 'Numbered') { $txtFind.Text } else { '' }
    $filterMode = if ($rbExclude.Checked) { 'Exclude' } else { 'Include' }

    $script:Plan = @(Build-Plan -Folder $FolderPath -Mode (Get-SelectedMode) `
        -NameMode $nameMode -Prefix $txtPrefix.Text -Suffix $txtSuffix.Text `
        -StartNumber ([int]$numStart.Value) -Digits ([int]$numDigits.Value) -SortBy $sortBy `
        -Recurse:$chkRecurse.Checked -ExtensionFilter $txtExtFilter.Text -FilterMode $filterMode `
        -ReplaceFrom $replaceFrom -ReplaceTo $txtReplaceNew.Text -CaseSensitive:$chkCase.Checked)

    $gridSeq.DataSource = ConvertTo-PreviewTable -Plan $script:Plan
    # 序号/类型/状态列收窄，把空间让给名称列
    $weights = @(8, 8, 18, 28, 28, 10)
    for ($i = 0; $i -lt $gridSeq.Columns.Count -and $i -lt $weights.Count; $i++) {
        $gridSeq.Columns[$i].FillWeight = $weights[$i]
    }
    foreach ($row in $gridSeq.Rows) {
        $status = [string]$row.Cells[5].Value
        if ($script:BlockedStatuses -contains $status) {
            $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Danger
            $row.DefaultCellStyle.BackColor = Get-Color $script:Ui.DangerB
        } elseif ($status -eq '未变化') {
            $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Muted
            $row.DefaultCellStyle.BackColor = Get-Color '#F8FAFC'
        } else {
            $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Text
        }
    }

    $blockedCount = @($script:Plan | Where-Object { $script:BlockedStatuses -contains $_.Status }).Count
    $unchangedCount = @($script:Plan | Where-Object Status -eq '未变化').Count
    $executableCount = @($script:Plan | Where-Object Status -eq '可执行').Count
    Update-SequenceStats -Total $script:Plan.Count -Executable $executableCount -Unchanged $unchangedCount -Blocked $blockedCount
    $script:SeqPlanDirty = $false
    Update-SequenceExecuteState

    # 位数不足时提示，避免出现 999 与 1000 混排
    $lblSeqCount.Text = "共 $($script:Plan.Count) 项"
    if ($nameMode -eq 'Numbered' -or $nameMode -eq 'PrefixNumberOriginal') {
        $maxText = (([int]$numStart.Value) + $script:Plan.Count - 1).ToString()
        if ($maxText.Length -gt [int]$numDigits.Value) {
            $lblSeqCount.Text = "共 $($script:Plan.Count) 项（注意：最大编号 $maxText 超过 $([int]$numDigits.Value) 位）"
        }
    }
    if ($blockedCount -gt 0) {
        Set-SeqHint -Text '存在冲突，无法执行' -Kind 'Warn'
    } elseif ($executableCount -gt 0) {
        Set-SeqHint -Text '已预览，未执行' -Kind 'Info'
    } else {
        Set-SeqHint -Text '没有可执行项' -Kind 'Warn'
    }
    return $script:Plan
}

function Load-SegmentFiles {
    param([string]$FolderPath)
    $script:FileItems = @(Get-SegmentFolderItems -FolderPath $FolderPath)
    $script:SegmentFolder = $FolderPath
}

function Rebuild-FixedControls {
    param([switch]$PreserveValues)
    # 窗口变宽/变窄后需要按容器实际宽度重新排布，先把已填的值缓存下来
    $savedTexts = @()
    $savedChecks = @()
    if ($PreserveValues) {
        foreach ($box in $script:FixedBoxes) { $savedTexts += $box.Text }
        foreach ($chk in $script:FixedChecks) { $savedChecks += $chk.Checked }
    }

    $pnlFixedValues.Controls.Clear()
    $script:FixedChecks = @()
    $script:FixedBoxes = @()
    $script:RestoringFixed = $true
    try {
        # 左栏放不下 6 段单行时折行，避免第 5/6 段输入框被裁掉
        $containerWidth = $pnlFixedValues.ClientSize.Width
        if ($containerWidth -le 0) { $containerWidth = $segBodyWidth }
        $minSlot = 144
        $cols = [math]::Max(1, [math]::Floor($containerWidth / $minSlot))
        if ($cols -gt $script:SegmentCount) { $cols = $script:SegmentCount }
        $slotWidth = [math]::Floor($containerWidth / $cols)
        $rowCount = [math]::Ceiling($script:SegmentCount / $cols)
        $rowHeight = 32
        $boxWidth = $slotWidth - 108
        if ($boxWidth -lt 36) { $boxWidth = 36 }
        $pnlFixedValues.Height = (4 + ($rowCount * $rowHeight))
        if ($cardFixed -and $cardFixed.Border) {
            $cardFixed.Border.Height = (45 + 8 + $pnlFixedValues.Height + 8)
        }

        for ($index = 1; $index -le $script:SegmentCount; $index++) {
            $slot = $index
            $col = ($index - 1) % $cols
            $row = [math]::Floor(($index - 1) / $cols)
            $offset = $col * $slotWidth
            $top = 4 + ($row * $rowHeight)

            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = "段$index"
            $lbl.Location = New-Object System.Drawing.Point($offset, ($top + 2))
            $lbl.Size = New-Object System.Drawing.Size(36, 22)
            $lbl.BackColor = [System.Drawing.Color]::Transparent
            $lbl.ForeColor = Get-Color $script:Ui.Muted
            $lbl.Font = $fontSmall
            $pnlFixedValues.Controls.Add($lbl)

            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = '固定'
            $chk.Location = New-Object System.Drawing.Point(($offset + 40), ($top + 1))
            $chk.Size = New-Object System.Drawing.Size(52, 24)
            $chk.BackColor = [System.Drawing.Color]::Transparent
            $chk.ForeColor = Get-Color $script:Ui.Text
            $chk.Tag = $slot
            $pnlFixedValues.Controls.Add($chk)

            $box = New-Object System.Windows.Forms.TextBox
            $box.Location = New-Object System.Drawing.Point(($offset + 98), $top)
            $box.Size = New-Object System.Drawing.Size($boxWidth, 26)
            $box.BorderStyle = 'FixedSingle'
            $box.Font = $fontInput
            $box.BackColor = Get-Color '#F1F5F9'
            $box.Tag = $slot
            $pnlFixedValues.Controls.Add($box)

            # 勾选后才让输入框变白，视觉上直接区分「固定/未固定」；用 Tag 传槽位，避免闭包作用域问题
            $chk.Add_CheckedChanged({
                $slotIndex = [int]$this.Tag
                $boxControl = $script:FixedBoxes[$slotIndex - 1]
                if ($boxControl) {
                    if ($this.Checked) {
                        $boxControl.BackColor = [System.Drawing.Color]::White
                    } else {
                        $boxControl.BackColor = Get-Color '#F1F5F9'
                    }
                }
                Apply-FixedToColumn -Index $slotIndex
            })
            $box.Add_TextChanged({ Apply-FixedToColumn -Index ([int]$this.Tag) })

            $script:FixedChecks += $chk
            $script:FixedBoxes += $box
        }

        if ($PreserveValues) {
            for ($index = 0; $index -lt $script:FixedBoxes.Count; $index++) {
                if ($index -lt $savedTexts.Count) { $script:FixedBoxes[$index].Text = $savedTexts[$index] }
                if ($index -lt $savedChecks.Count) { $script:FixedChecks[$index].Checked = $savedChecks[$index] }
            }
        }
    } finally {
        $script:RestoringFixed = $false
    }
    $script:FixedPanelWidth = $pnlFixedValues.ClientSize.Width
}

function Apply-FixedToColumn {
    # 思路：勾选固定后把该值刷到所有行的这一列；不勾选则保留用户手改的内容
    param([int]$Index)
    # 重排固定值控件时回填值也会触发事件，此时不要回写表格
    if ($script:RestoringFixed) { return }
    $position = $Index - 1
    if ($position -lt 0 -or $position -ge $script:FixedChecks.Count) { return }
    if (-not $script:FixedChecks[$position].Checked) { return }
    if ($gridSeg.Columns.Count -le $Index) { return }
    $value = $script:FixedBoxes[$position].Text
    $script:SuppressRefresh = $true
    try {
        foreach ($gridRow in $gridSeg.Rows) {
            # 列 0 是文件名，段i 在第 i 列（不是 i+1），写错会整体串到下一段
            $gridRow.Cells[$Index].Value = $value
        }
    } finally {
        $script:SuppressRefresh = $false
    }
    try { [void](Refresh-SegmentStatus) } catch { $lblSegCount.Text = "刷新失败：$($_.Exception.Message)" }
}

function Rebuild-SegmentGrid {
    $script:SegmentCount = [int]$numCount.Value
    $script:SuppressRefresh = $true
    try {
        Rebuild-FixedControls
        $gridSeg.Columns.Clear()
        $gridSeg.Rows.Clear()

        $nameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $nameColumn.HeaderText = '文件名'
        $nameColumn.Width = 200
        $nameColumn.ReadOnly = $true
        $nameColumn.DefaultCellStyle.BackColor = Get-Color '#F8F9FB'
        [void]$gridSeg.Columns.Add($nameColumn)

        # 固定列占 200/210/95，剩余宽度按段数均分，段数多时也不出现横向滚动条
        $gridWidth = $gridSeg.ClientSize.Width
        if ($gridWidth -le 0) { $gridWidth = $segmentBodyWidth }
        $segmentWidth = [math]::Max(56, [math]::Floor(($gridWidth - 200 - 210 - 95 - 10) / $script:SegmentCount))
        for ($segmentIndex = 1; $segmentIndex -le $script:SegmentCount; $segmentIndex++) {
            $segmentColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $segmentColumn.HeaderText = "段$segmentIndex"
            $segmentColumn.Width = $segmentWidth
            $segmentColumn.DefaultCellStyle.BackColor = Get-Color '#FFFBEB'
            [void]$gridSeg.Columns.Add($segmentColumn)
        }

        $newNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $newNameColumn.HeaderText = '新文件名'
        $newNameColumn.Width = 210
        $newNameColumn.ReadOnly = $true
        $newNameColumn.DefaultCellStyle.BackColor = Get-Color '#F8F9FB'
        [void]$gridSeg.Columns.Add($newNameColumn)

        $statusColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $statusColumn.HeaderText = '状态'
        $statusColumn.Width = 95
        $statusColumn.ReadOnly = $true
        [void]$gridSeg.Columns.Add($statusColumn)

        foreach ($item in $script:FileItems) {
            $rowIndex = $gridSeg.Rows.Add()
            $row = $gridSeg.Rows[$rowIndex]
            $row.Cells[0].Value = $item.Name
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
            $segments = Split-NameToSegments -BaseName $baseName -Separator $txtSeparator.Text -SegmentCount $script:SegmentCount
            for ($segmentIndex = 0; $segmentIndex -lt $script:SegmentCount; $segmentIndex++) {
                $value = $segments[$segmentIndex]
                if ($script:FixedChecks[$segmentIndex].Checked -and -not [string]::IsNullOrWhiteSpace($script:FixedBoxes[$segmentIndex].Text)) {
                    $value = $script:FixedBoxes[$segmentIndex].Text
                }
                $row.Cells[$segmentIndex + 1].Value = $value
            }
            $row.Cells[$script:SegmentCount + 1].Value = ''
            $row.Cells[$script:SegmentCount + 2].Value = ''
        }
    } finally {
        $script:SuppressRefresh = $false
    }
}

function Update-SegmentStats {
    param([int]$Total, [int]$Executable, [int]$Unchanged, [int]$Blocked)
    $lblSegCount.Text = "共 $Total 个文件"
    Set-StatTag -Label $lblSegExec -Text "可执行 $Executable" -Kind $(if ($Executable -gt 0) { 'Success' } else { 'Muted' })
    Set-StatTag -Label $lblSegUnchanged -Text "未变化 $Unchanged" -Kind 'Muted'
    Set-StatTag -Label $lblSegBlocked -Text "有问题 $Blocked" -Kind $(if ($Blocked -gt 0) { 'Danger' } else { 'Muted' })
}

function Refresh-SegmentStatus {
    # 思路：把表格内容转成计划，回填「新文件名」「状态」，并统计各类数量
    if ($gridSeg.Rows.Count -eq 0) {
        $lblSegCount.Text = '当前文件夹没有可处理的文件。'
        Update-SegmentStats -Total 0 -Executable 0 -Unchanged 0 -Blocked 0
        return @()
    }
    $script:SuppressRefresh = $true
    try {
        $rows = @(Get-RowRecords -Grid $gridSeg -SegmentCount $script:SegmentCount -FolderPath $script:SegmentFolder)
        $plan = @(Build-ComposerPlan -Rows $rows -Separator $txtSeparator.Text)

        for ($index = 0; $index -lt $plan.Count; $index++) {
            $row = $gridSeg.Rows[$index]
            $row.Cells[$script:SegmentCount + 1].Value = $plan[$index].NewName
            $row.Cells[$script:SegmentCount + 2].Value = $plan[$index].Status
            if ($script:BlockedStatuses -contains $plan[$index].Status) {
                $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Danger
                $row.DefaultCellStyle.BackColor = Get-Color $script:Ui.DangerB
            } elseif ($plan[$index].Status -eq '未变化') {
                $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Muted
                $row.DefaultCellStyle.BackColor = Get-Color '#F8FAFC'
            } else {
                $row.DefaultCellStyle.ForeColor = Get-Color $script:Ui.Text
            }
        }

        $executable = @($plan | Where-Object Status -eq '可执行').Count
        $unchanged = @($plan | Where-Object Status -eq '未变化').Count
        $blocked = @($plan | Where-Object { $script:BlockedStatuses -contains $_.Status }).Count
        Update-SegmentStats -Total $plan.Count -Executable $executable -Unchanged $unchanged -Blocked $blocked
        return $plan
    } finally {
        $script:SuppressRefresh = $false
    }
}

# ===== 事件绑定 =====

Show-View 'sequence'

$btnSeqBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '选择要批量重命名的文件夹'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSeqFolder.Text = $dialog.SelectedPath
        Update-SequenceSample
        Set-SequencePlanDirty
    }
})

$txtPrefix.Add_TextChanged({ Update-SequenceSample; Set-SequencePlanDirty })
$txtSuffix.Add_TextChanged({ Update-SequenceSample; Set-SequencePlanDirty })
$numStart.Add_ValueChanged({ Update-SequenceSample; Set-SequencePlanDirty })
$numDigits.Add_ValueChanged({ Update-SequenceSample; Set-SequencePlanDirty })
$cmbNameMode.Add_SelectedIndexChanged({ Update-ReplaceAvailability; Update-SequenceSample; Set-SequencePlanDirty })
$chkReplace.Add_CheckedChanged({ Update-SequenceSample; Set-SequencePlanDirty })
$txtSeqFolder.Add_TextChanged({ Set-SequencePlanDirty })
$cmbSort.Add_SelectedIndexChanged({ Set-SequencePlanDirty })
$txtExtFilter.Add_TextChanged({ Set-SequencePlanDirty })
$txtFind.Add_TextChanged({ Set-SequencePlanDirty })
$txtReplaceNew.Add_TextChanged({ Set-SequencePlanDirty })
$chkCase.Add_CheckedChanged({ Set-SequencePlanDirty })
$rbFiles.Add_CheckedChanged({ Set-SequencePlanDirty })
$rbFolders.Add_CheckedChanged({ Set-SequencePlanDirty })
$rbBoth.Add_CheckedChanged({ Set-SequencePlanDirty })
$rbInclude.Add_CheckedChanged({ Set-SequencePlanDirty })
$rbExclude.Add_CheckedChanged({ Set-SequencePlanDirty })

# 递归只处理文件：同时禁用「文件夹」和「两者」，避免 UI 可选但实际忽略
$chkRecurse.Add_CheckedChanged({
    if ($chkRecurse.Checked) {
        $rbFiles.Checked = $true
        $rbFolders.Enabled = $false
        $rbBoth.Enabled = $false
        $lblRecurseHint.Text = '已开启：只处理文件，文件夹/两者已禁用'
        $lblRecurseHint.ForeColor = Get-Color $script:Ui.Warn
    } else {
        $rbFolders.Enabled = $true
        $rbBoth.Enabled = $true
        $lblRecurseHint.Text = '递归开启后只处理文件，文件夹选项会禁用'
        $lblRecurseHint.ForeColor = Get-Color $script:Ui.Muted
    }
    Set-SequencePlanDirty
})

$btnSeqPreview.Add_Click({
    try {
        [void](Update-SequencePreview -FolderPath $txtSeqFolder.Text.Trim())
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '预览失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnSeqExecute.Add_Click({
    try {
        if ($script:Plan.Count -eq 0) { throw '请先点击「预览」生成重命名计划。' }
        $blocked = @($script:Plan | Where-Object { $script:BlockedStatuses -contains $_.Status })
        if ($blocked.Count -gt 0) {
            throw "有 $($blocked.Count) 条记录存在问题（红色行），请调整规则后重新预览。"
        }
        $validCount = @($script:Plan | Where-Object Status -eq '可执行').Count
        if ($validCount -eq 0) { throw '没有需要重命名的对象。' }

        $fileCount = @($script:Plan | Where-Object { $_.Type -eq '文件' -and $_.Status -eq '可执行' }).Count
        $dirCount = @($script:Plan | Where-Object { $_.Type -eq '文件夹' -and $_.Status -eq '可执行' }).Count
        $samples = @($script:Plan | Where-Object Status -eq '可执行' | Select-Object -First 5 | ForEach-Object { "$($_.SourceName) -> $($_.NewName)" })
        $confirmText = "目标目录：$($txtSeqFolder.Text.Trim())`r`n计划 $($script:Plan.Count) 项，将变更 $validCount 项（文件 $fileCount / 文件夹 $dirCount）`r`n计划状态：当前预览（未过期）`r`n`r`n示例：`r`n" + ($samples -join "`r`n") + "`r`n`r`n执行后可用撤销回退。是否继续？"
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmText,
            '确认执行', [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $result = Invoke-RenamePlan -Plan $script:Plan
        $message = Format-RenameResultMessage -Result $result -Action '执行'
        $icon = if ($result.NeedsAttentionCount -gt 0 -or $result.Outcome -eq 'CompleteFailure') {
            [System.Windows.Forms.MessageBoxIcon]::Warning
        } else { [System.Windows.Forms.MessageBoxIcon]::Information }
        [void][System.Windows.Forms.MessageBox]::Show($message, '执行结果',
            [System.Windows.Forms.MessageBoxButtons]::OK, $icon)

        Notify-DirectoryChanged
        $lblSeqCount.Text = '执行完成，请重新预览。'
        Set-SeqHint -Text '执行完成' -Kind 'Ok'
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '执行失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnSeqUndo.Add_Click({
    try {
        $logs = @(Get-UndoLogs)
        if ($logs.Count -eq 0) { throw '没有找到可撤销的操作记录。' }
        $logFile = $logs[0]
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "将撤销最近一次操作（$($logFile.LastWriteTime.ToString('MM-dd HH:mm:ss'))）。`r`n当前共 $($logs.Count) 步可撤销，撤销后还可继续回退更早的操作。是否继续？",
            '确认撤销', [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $result = Invoke-UndoLatest
        $message = "撤销结果：$($result.Outcome)`r`n已还原 $($result.RestoredCount) 个，失败 $($result.Errors.Count) 个。"
        if ($result.NeedsAttentionCount -gt 0) {
            $message += "`r`n待恢复 $($result.NeedsAttentionCount) 个（未覆盖占用位置）。"
        }
        if ($result.Errors.Count -gt 0) {
            $message += "`r`n`r`n前 10 条错误：`r`n" + (($result.Errors | Select-Object -First 10) -join "`r`n")
            $message += "`r`n`r`n该条记录已保留，处理完占用后可再次点击撤销重试。请勿删除 ~rn_ 临时名。"
        } else {
            $message += "`r`n`r`n该步已标记完成，可继续撤销更早的操作。"
        }
        [void][System.Windows.Forms.MessageBox]::Show($message, '撤销结果',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Notify-DirectoryChanged
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '撤销失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnSegBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $current = $txtSegFolder.Text.Trim()
    if ($current -and (Test-Path -LiteralPath $current)) { $dialog.SelectedPath = $current }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSegFolder.Text = $dialog.SelectedPath
        try {
            Load-SegmentFiles -FolderPath $dialog.SelectedPath
            Rebuild-SegmentGrid
            [void](Refresh-SegmentStatus)
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '读取失败',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})

$numCount.Add_ValueChanged({
    if ($script:FileItems.Count -gt 0) {
        Rebuild-SegmentGrid
        [void](Refresh-SegmentStatus)
    } else {
        $script:SegmentCount = [int]$numCount.Value
        Rebuild-FixedControls
    }
})

$btnSplit.Add_Click({
    try {
        if ($script:FileItems.Count -eq 0) { throw '当前文件夹没有文件可拆分。' }
        $separator = $txtSeparator.Text
        $script:SuppressRefresh = $true
        try {
            foreach ($gridRow in $gridSeg.Rows) {
                $name = [string]$gridRow.Cells[0].Value
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)
                $segments = Split-NameToSegments -BaseName $baseName -Separator $separator -SegmentCount $script:SegmentCount
                for ($segmentIndex = 0; $segmentIndex -lt $script:SegmentCount; $segmentIndex++) {
                    $value = $segments[$segmentIndex]
                    if ($script:FixedChecks[$segmentIndex].Checked -and -not [string]::IsNullOrWhiteSpace($script:FixedBoxes[$segmentIndex].Text)) {
                        $value = $script:FixedBoxes[$segmentIndex].Text
                    }
                    $gridRow.Cells[$segmentIndex + 1].Value = $value
                }
            }
        } finally {
            $script:SuppressRefresh = $false
        }
        [void](Refresh-SegmentStatus)
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '拆分失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnClearSegments.Add_Click({
    $script:SuppressRefresh = $true
    try {
        foreach ($gridRow in $gridSeg.Rows) {
            for ($segmentIndex = 0; $segmentIndex -lt $script:SegmentCount; $segmentIndex++) {
                if ($script:FixedChecks[$segmentIndex].Checked -and -not [string]::IsNullOrWhiteSpace($script:FixedBoxes[$segmentIndex].Text)) {
                    $gridRow.Cells[$segmentIndex + 1].Value = $script:FixedBoxes[$segmentIndex].Text
                } else {
                    $gridRow.Cells[$segmentIndex + 1].Value = ''
                }
            }
        }
    } finally {
        $script:SuppressRefresh = $false
    }
    [void](Refresh-SegmentStatus)
})

$btnSegRefresh.Add_Click({
    try {
        Load-SegmentFiles -FolderPath $txtSegFolder.Text.Trim()
        Rebuild-SegmentGrid
        [void](Refresh-SegmentStatus)
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '读取失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$gridSeg.Add_CellValueChanged({
    if ($script:SuppressRefresh) { return }
    try { [void](Refresh-SegmentStatus) } catch {
        $lblSegCount.Text = "刷新失败：$($_.Exception.Message)"
    }
})

$gridSeg.Add_CurrentCellDirtyStateChanged({
    if ($gridSeg.IsCurrentCellDirty) { $gridSeg.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
})

$btnSegRun.Add_Click({
    try {
        $plan = @(Refresh-SegmentStatus)
        if ($plan.Count -eq 0) { throw '没有可处理的文件。' }
        $blocked = @($plan | Where-Object { $script:BlockedStatuses -contains $_.Status })
        if ($blocked.Count -gt 0) {
            throw "有 $($blocked.Count) 行存在问题（红色行），请修正后再执行。"
        }
        $executable = @($plan | Where-Object Status -eq '可执行')
        if ($executable.Count -eq 0) { throw '没有需要改名的文件（全部未变化）。' }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "即将改名 $($executable.Count) 个文件，是否继续？`r`n改名后可在「序号编号」页点「撤销」回退。",
            '确认执行', [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $result = Invoke-ComposerRename -Plan $executable
        $message = Format-RenameResultMessage -Result $result -Action '分段改名'
        $icon = if ($result.NeedsAttentionCount -gt 0 -or $result.Outcome -eq 'CompleteFailure') {
            [System.Windows.Forms.MessageBoxIcon]::Warning
        } else { [System.Windows.Forms.MessageBoxIcon]::Information }
        [void][System.Windows.Forms.MessageBox]::Show($message, '执行结果',
            [System.Windows.Forms.MessageBoxButtons]::OK, $icon)

        Notify-DirectoryChanged
        Load-SegmentFiles -FolderPath $txtSegFolder.Text.Trim()
        Rebuild-SegmentGrid
        [void](Refresh-SegmentStatus)
        $script:SegmentStale = $false
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '无法执行',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$form.Add_Shown({
    Update-ReplaceAvailability
    Update-SequenceSample
    Update-UndoButton
    # 段数默认 3 时先建一次固定值控件，避免首次进入分段视图时是空的
    Rebuild-FixedControls
    $pending = @(Get-PendingBatchRenamerJournals)
    if ($pending.Count -gt 0) {
        $summary = Format-BatchRenamerPendingSummary -Journals $pending
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "发现 $($pending.Count) 个未完成事务，不会自动删除 ~rn_ 临时名。`r`n`r`n$summary`r`n`r`n是否现在尝试安全恢复？",
            '未完成事务',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
            $recoverLines = New-Object System.Collections.ArrayList
            foreach ($journal in $pending) {
                try {
                    $recoverResult = Invoke-BatchRenamerRecoverJournal -Journal $journal
                    [void]$recoverLines.Add(('事务 {0}：{1}' -f $journal.TransactionId, $recoverResult.Outcome))
                    if ($recoverResult.Errors.Count -gt 0) {
                        [void]$recoverLines.Add((($recoverResult.Errors | Select-Object -First 5) -join "`r`n"))
                    }
                } catch {
                    [void]$recoverLines.Add(('事务 {0} 恢复失败：{1}' -f $journal.TransactionId, $_.Exception.Message))
                }
            }
            [void][System.Windows.Forms.MessageBox]::Show(($recoverLines -join "`r`n"), '恢复结果',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Update-UndoButton
        }
    }
    # 启动参数带入路径时直接出一份预览，省掉一次手动点击
    if (-not [string]::IsNullOrWhiteSpace($txtSeqFolder.Text)) {
        try {
            [void](Update-SequencePreview -FolderPath $txtSeqFolder.Text.Trim())
        } catch {
            $lblSeqCount.Text = "预览失败：$($_.Exception.Message)"
        }
    }
    try { $form.Activate() } catch { }
})

# 窗口尺寸变化后按新宽度重排固定值槽位，否则缩窄时最后一段会被裁掉
$form.Add_Resize({
    if ($script:FixedChecks.Count -eq 0) { return }
    if ([math]::Abs($pnlFixedValues.ClientSize.Width - $script:FixedPanelWidth) -lt 8) { return }
    Rebuild-FixedControls -PreserveValues
})

# 回到窗口时刷新可撤销步数（改名可能发生在分段视图里）
$form.Add_Activated({ Update-UndoButton })

# ---------------------------------------------------------------------------
# 诊断：构建界面但不显示，验证两条链路
# ---------------------------------------------------------------------------
if ($Diagnose) {
    try {
        if ([string]::IsNullOrWhiteSpace($Folder)) {
            Write-Output 'DIAG_SKIP=FolderMissing'
            throw '诊断需要有效的 -Folder 目录，未提供目录不能计为通过。'
        }
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
            Write-Output 'DIAG_SKIP=FolderInvalid'
            throw "诊断目录不存在或不是文件夹：$Folder"
        }
        if (-not [string]::IsNullOrWhiteSpace($Folder)) {
            $txtSeqFolder.Text = $Folder
            $seqPlan = @(Update-SequencePreview -FolderPath $Folder)
            Write-Output ("DIAG_SEQ_ROWS={0} COLS={1} PLAN={2}" -f $gridSeq.Rows.Count, $gridSeq.Columns.Count, $seqPlan.Count)

            $txtSegFolder.Text = $Folder
            Load-SegmentFiles -FolderPath $Folder
            Rebuild-SegmentGrid
            $segPlan = @(Refresh-SegmentStatus)
            Write-Output ("DIAG_SEG_ROWS={0} COLS={1} PLAN={2}" -f $gridSeg.Rows.Count, $gridSeg.Columns.Count, $segPlan.Count)
            foreach ($gridRow in $gridSeg.Rows) {
                $cells = @()
                foreach ($cell in $gridRow.Cells) { $cells += [string]$cell.Value }
                Write-Output ('  ' + ($cells -join ' | '))
            }

            # 回归检查：段1 的固定值必须落在段1 列，串到下一列就说明列索引错位
            if ($gridSeg.Rows.Count -gt 0 -and $script:FixedChecks.Count -ge 2) {
                $lastBox = $script:FixedBoxes[$script:FixedBoxes.Count - 1]
                $lastRight = $lastBox.Left + $lastBox.Width
                Write-Output ('GEOM_FORM=' + $form.ClientSize.Width + ' CARD=' + $cardWidth + ' SEG_BODY=' + $segmentBodyWidth + ' FIXED_PANEL=' + $pnlFixedValues.ClientSize.Width + ' SEGMENTS=' + $script:SegmentCount + ' BOX1=' + $script:FixedBoxes[0].Width + ' LAST_RIGHT=' + $lastRight)
                # 最后一段的输入框超出容器就说明被裁切了
                if ($lastRight -gt $pnlFixedValues.ClientSize.Width) {
                    throw "固定值行溢出：最后一段右边界 $lastRight 超过容器 $($pnlFixedValues.ClientSize.Width)"
                }
                # 查找替换行：启用复选框不能与「查找」标签重叠
                Write-Output ('REPLACE_ROW_CHK=' + $chkReplace.Right + ' FIND=' + $lblFind.Left + ' HINT=' + $lblReplaceHint.Left + '/' + $lblReplaceHint.Width)
                if ($chkReplace.Right -gt $lblFind.Left) { throw '查找替换行排版重叠：启用复选框压住了查找标签' }
                $script:FixedBoxes[0].Text = '固定甲'
                $script:FixedChecks[0].Checked = $true
                $landedValue = [string]$gridSeg.Rows[0].Cells[1].Value
                Write-Output ("FIXED_CHECK_CELL1={0}" -f $landedValue)
                if ($landedValue -ne '固定甲') { throw "固定值列错位：段1 列实际为 $landedValue" }

                $savedCount = $script:SegmentCount
                $script:SegmentCount = 6
                Rebuild-FixedControls
                $maxRight = 0
                foreach ($fixedBox in $script:FixedBoxes) {
                    $right = $fixedBox.Left + $fixedBox.Width
                    if ($right -gt $maxRight) { $maxRight = $right }
                }
                Write-Output ('GEOM_FIXED6 COUNT=' + $script:FixedBoxes.Count + ' MAX_RIGHT=' + $maxRight + ' PANEL=' + $pnlFixedValues.ClientSize.Width)
                if ($script:FixedBoxes.Count -ne 6) { throw '固定值应为 6 段' }
                if ($maxRight -gt $pnlFixedValues.ClientSize.Width) {
                    throw "6 段固定值溢出：右边界 $maxRight 超过容器 $($pnlFixedValues.ClientSize.Width)"
                }
                $script:SegmentCount = $savedCount
                Rebuild-FixedControls -PreserveValues
            }
        }

        Show-View 'segment'
        # 窗体未显示时子控件的 Visible 恒为 false，这里断言当前视图状态变量
        Write-Output ("ACTIVE_VIEW={0}" -f $script:ActiveView)
        Write-Output 'DIAGNOSE=PASS'
    } catch {
        Write-Output 'DIAGNOSE=FAIL'
        Write-Output ('ERROR: ' + $_.Exception.Message)
        exit 1
    }
    exit 0
}

try {
    [void]$form.ShowDialog()
} catch {
    $logPath = Join-Path $env:TEMP 'BatchRenamer_error.log'
    "时间：$(Get-Date -Format s)`r`n错误：$($_.Exception.Message)`r`n堆栈：$($_.ScriptStackTrace)" |
        Set-Content -LiteralPath $logPath -Encoding UTF8
    [void][System.Windows.Forms.MessageBox]::Show("程序启动失败：$($_.Exception.Message)`r`n日志：$logPath", '错误',
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
