#requires -Version 5.1
<#
分段拼接改名
- 把文件名拆成若干段，在表格里逐行查看/修改，再按连接符重新拼成新文件名
- 每段可勾选「固定」：自动填入该值，但单元格仍可手改（适合「编辑姓名固定、个别要改」）
- 分隔符不规律时可在表格里直接改，不依赖自动拆分
- 执行采用两阶段改名，写 undo 日志，可回主界面「撤销」
#>
param(
    [string]$Folder = '',
    [switch]$SelfTest,
    [switch]$Diagnose
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:CorePath = Join-Path $PSScriptRoot 'BatchRenamer.Core.ps1'
if (-not (Test-Path -LiteralPath $script:CorePath)) {
    throw '找不到 BatchRenamer.Core.ps1，请与本脚本放在同一文件夹。'
}
. $script:CorePath

$script:LogFolder = Join-Path $env:TEMP 'BatchRenamerLogs'
$script:FileItems = @()
$script:FixedChecks = @()
$script:FixedBoxes = @()
$script:SegmentCount = 3
$script:BlockedStatuses = @('无效字符', '名称重复', '目标已存在', '名称为空')
# 程序内部批量刷新单元格时会触发 CellValueChanged，用该开关阻断递归刷新
$script:SuppressRefresh = $false
$script:NaturalComparer = [System.Globalization.CultureInfo]::CurrentCulture.CompareInfo

# ---------------------------------------------------------------------------
# 纯逻辑函数
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

function Test-NameValid {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalidChars) {
        if ($Name.IndexOf($char) -ge 0) { return $false }
    }
    if ($Name.EndsWith('.') -or $Name.EndsWith(' ')) { return $false }
    $reserved = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5',
                  'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4',
                  'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    $basePart = $Name.Split('.')[0].ToUpperInvariant()
    if ($reserved -contains $basePart) { return $false }
    return $true
}

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
    # 思路：一次性算出每行的新名称与状态，冲突检测覆盖跨行重名和磁盘已存在
    param(
        [object[]]$Rows,
        [string]$Separator
    )
    $plan = New-Object System.Collections.ArrayList
    $usedNames = @{ }

    foreach ($row in $Rows) {
        $segments = [string[]]$row.Segments
        $newName = Compose-Name -Segments $segments -Separator $Separator -Extension $row.Extension
        $status = '可执行'

        $hasContent = @($segments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        # 空段保留连接符，例如 2 + 空 + 1 => 2--1.docx；只有全部为空才拦
        if ($hasContent -eq 0 -or [string]::IsNullOrWhiteSpace($newName)) {
            $status = '名称为空'
        } elseif (-not (Test-NameValid -Name $newName)) {
            $status = '无效字符'
        } elseif ([string]::Equals($row.Name, $newName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $status = '未变化'
        } elseif ($usedNames.ContainsKey($newName.ToLowerInvariant())) {
            $status = '名称重复'
        } elseif (Test-Path -LiteralPath (Join-Path $row.ParentDir $newName)) {
            $status = '目标已存在'
        }

        if (-not [string]::IsNullOrWhiteSpace($newName) -and -not $usedNames.ContainsKey($newName.ToLowerInvariant())) {
            $usedNames[$newName.ToLowerInvariant()] = $true
        }

        [void]$plan.Add([pscustomobject]@{
            SourcePath = $row.SourcePath
            ParentDir  = $row.ParentDir
            SourceName = $row.Name
            NewName    = $newName
            Status     = $status
            Segments   = $segments
        })
    }
    return $plan.ToArray()
}

# Invoke-ComposerRename 由 BatchRenamer.Core.ps1 提供。

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
        [void]$rows.Add([pscustomobject]@{
            SourcePath = (Join-Path $FolderPath $name)
            ParentDir  = $FolderPath
            Name       = $name
            Extension  = $extension
            Segments   = [string[]]$segments.ToArray()
        })
    }
    return $rows.ToArray()
}

function Invoke-SelfTest {
    $testRoot = Join-Path $env:TEMP ('ComposerSelfTest_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $logFolder = Join-Path $testRoot 'logs'
    $root = Join-Path $testRoot 'data'
    [void](New-Item -ItemType Directory -Path $root -Force)
    [void](New-Item -ItemType Directory -Path $logFolder -Force)
    $officialBefore = Get-BatchRenamerLogFingerprint -LogFolder (Get-BatchRenamerLogFolder)
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

        $result = Invoke-ComposerRename -Plan $plan -LogFolder $logFolder
        Write-Output ("RENAMED={0} ERRORS={1}" -f $result.SuccessCount, $result.Errors.Count)
        if ($result.SuccessCount -ne 2) { throw '自测失败：改名数量不符' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '2025-张三-李四.docx'))) { throw '自测失败：缺少目标文件1' }
        if (-not (Test-Path -LiteralPath (Join-Path $root '甲X乙-张三-李四.txt'))) { throw '自测失败：缺少目标文件2' }
        if (-not $result.LogPath -or -not (Test-Path -LiteralPath $result.LogPath)) { throw '自测失败：未生成撤销日志' }
        $logRecord = @(Import-Csv -LiteralPath $result.LogPath -Encoding UTF8)[0]
        if (-not $logRecord.OldPath -or -not $logRecord.NewPath) { throw '自测失败：日志缺少 OldPath/NewPath' }
        $undo = Invoke-UndoLatest -LogFolder $logFolder
        if (-not $undo.MarkedDone) { throw '自测失败：分段撤销未完成' }
        $officialAfter = Get-BatchRenamerLogFingerprint -LogFolder (Get-BatchRenamerLogFolder)
        if ($officialBefore -ne $officialAfter) { throw '自测失败：正式日志被改动' }
        Write-Output 'SELFTEST=PASS'
    } finally {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($SelfTest) {
    try {
        Invoke-SelfTest
        exit 0
    } catch {
        Write-Output 'SELFTEST=FAIL'
        Write-Output ('ERROR: ' + $_.Exception.Message)
        exit 1
    }
}

# ---------------------------------------------------------------------------
# 界面
# ---------------------------------------------------------------------------
$font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$headerFont = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$form = New-Object System.Windows.Forms.Form
$form.Text = '分段拼接改名'
$form.ClientSize = New-Object System.Drawing.Size(920, 660)
$form.MinimumSize = New-Object System.Drawing.Size(940, 620)
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F7F9FC')

$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = '目标文件夹'
$lblFolder.Location = New-Object System.Drawing.Point(12, 16)
$lblFolder.Size = New-Object System.Drawing.Size(70, 22)
$lblFolder.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Text = $Folder
$txtFolder.Location = New-Object System.Drawing.Point(88, 13)
$txtFolder.Size = New-Object System.Drawing.Size(700, 26)
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '浏览...'
$btnBrowse.Location = New-Object System.Drawing.Point(798, 11)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($btnBrowse)

$grpRule = New-Object System.Windows.Forms.GroupBox
$grpRule.Text = '组合规则'
$grpRule.Location = New-Object System.Drawing.Point(12, 48)
$grpRule.Size = New-Object System.Drawing.Size(890, 64)
$grpRule.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F7F9FC')
$form.Controls.Add($grpRule)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = '文件名由几段组成'
$lblCount.Location = New-Object System.Drawing.Point(12, 26)
$lblCount.Size = New-Object System.Drawing.Size(110, 22)
$lblCount.BackColor = [System.Drawing.Color]::Transparent
$grpRule.Controls.Add($lblCount)

$numCount = New-Object System.Windows.Forms.NumericUpDown
$numCount.Minimum = 1
$numCount.Maximum = 6
$numCount.Value = 3
$numCount.Location = New-Object System.Drawing.Point(124, 24)
$numCount.Size = New-Object System.Drawing.Size(50, 26)
$grpRule.Controls.Add($numCount)

$lblSeparator = New-Object System.Windows.Forms.Label
$lblSeparator.Text = '连接符'
$lblSeparator.Location = New-Object System.Drawing.Point(196, 26)
$lblSeparator.Size = New-Object System.Drawing.Size(50, 22)
$lblSeparator.BackColor = [System.Drawing.Color]::Transparent
$grpRule.Controls.Add($lblSeparator)

$txtSeparator = New-Object System.Windows.Forms.TextBox
$txtSeparator.Text = '-'
$txtSeparator.Location = New-Object System.Drawing.Point(248, 23)
$txtSeparator.Size = New-Object System.Drawing.Size(60, 26)
$grpRule.Controls.Add($txtSeparator)

$btnSplit = New-Object System.Windows.Forms.Button
$btnSplit.Text = '按连接符拆分'
$btnSplit.Location = New-Object System.Drawing.Point(330, 22)
$btnSplit.Size = New-Object System.Drawing.Size(110, 28)
$grpRule.Controls.Add($btnSplit)

$btnClearSegments = New-Object System.Windows.Forms.Button
$btnClearSegments.Text = '清空段值'
$btnClearSegments.Location = New-Object System.Drawing.Point(448, 22)
$btnClearSegments.Size = New-Object System.Drawing.Size(90, 28)
$grpRule.Controls.Add($btnClearSegments)

$grpFixed = New-Object System.Windows.Forms.GroupBox
$grpFixed.Text = '分段固定值（勾选「固定」= 自动填入所有行；单元格仍可单独手改）'
$grpFixed.Location = New-Object System.Drawing.Point(12, 120)
$grpFixed.Size = New-Object System.Drawing.Size(890, 70)
$grpFixed.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F7F9FC')
$form.Controls.Add($grpFixed)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '请选择文件夹后填写表格。'
$lblStatus.Location = New-Object System.Drawing.Point(12, 198)
$lblStatus.Size = New-Object System.Drawing.Size(890, 22)
$lblStatus.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblStatus)

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 224)
$dgv.Size = New-Object System.Drawing.Size(890, 380)
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.AllowUserToResizeRows = $false
$dgv.RowHeadersVisible = $false
$dgv.SelectionMode = 'FullRowSelect'
$dgv.MultiSelect = $false
$dgv.EditMode = 'EditOnEnter'
$dgv.AutoSizeColumnsMode = 'None'
$dgv.BackgroundColor = [System.Drawing.Color]::White
$dgv.BorderStyle = 'FixedSingle'
$dgv.CellBorderStyle = 'SingleHorizontal'
$dgv.GridColor = [System.Drawing.ColorTranslator]::FromHtml('#E3E8EF')
$dgv.RowTemplate.Height = 30
$dgv.EnableHeadersVisualStyles = $false
$dgv.ColumnHeadersHeight = 34
$dgv.ColumnHeadersHeightSizeMode = 'DisableResizing'
$dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#E9EFF8')
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#1F2937')
$dgv.ColumnHeadersDefaultCellStyle.Font = $headerFont
$dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FBFCFE')
$dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml('#DCE9FB')
$dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.ColorTranslator]::FromHtml('#111827')
$dgv.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
$form.Controls.Add($dgv)

# 统一按钮观感：主按钮实心蓝，其余白底浅边框
function Initialize-ButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button,
        [switch]$Primary
    )
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($Primary) {
        $Button.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#1D4ED8')
        $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#3B82F6')
    } else {
        $Button.BackColor = [System.Drawing.Color]::White
        $Button.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#1F2937')
        $Button.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#CBD5E1')
        $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml('#F1F5F9')
    }
}

Initialize-ButtonStyle -Button $btnBrowse
Initialize-ButtonStyle -Button $btnSplit
Initialize-ButtonStyle -Button $btnClearSegments

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '执行改名'
$btnRun.Location = New-Object System.Drawing.Point(692, 614)
$btnRun.Size = New-Object System.Drawing.Size(120, 34)
$form.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = '关闭'
$btnClose.Location = New-Object System.Drawing.Point(822, 614)
$btnClose.Size = New-Object System.Drawing.Size(80, 34)
$form.Controls.Add($btnClose)

function Load-FileItems {
    # 思路：只取当前目录的普通文件，排除隐藏/系统文件，按自然序排列
    $folder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) { throw '请先选择目标文件夹。' }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw '目标文件夹不存在，请重新选择。' }

    $files = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction Stop |
        Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::Hidden) -and
                        -not ($_.Attributes -band [System.IO.FileAttributes]::System) })
    $script:FileItems = @($files | Sort-Object { Get-NaturalKey -Name $_.Name })
    $script:FolderPath = $folder
}

function Rebuild-FixedControls {
    $grpFixed.Controls.Clear()
    $script:FixedChecks = @()
    $script:FixedBoxes = @()
    # 槽宽按段数自适应，避免 4 段以上控件跑到分组框外面
    # 单一档位文字（36）+ 勾选（52）+ 输入框，输入框宽度要再留出右侧间隙，否则会压到下一段的标签
    $slotWidth = [math]::Max(130, [math]::Floor(866 / $script:SegmentCount))
    $boxWidth = [math]::Max(44, $slotWidth - 108)
    for ($index = 1; $index -le $script:SegmentCount; $index++) {
        $slot = $index
        $offset = 12 + (($index - 1) * $slotWidth)

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "段$index"
        $lbl.Location = New-Object System.Drawing.Point($offset, 30)
        $lbl.Size = New-Object System.Drawing.Size(36, 22)
        $lbl.BackColor = [System.Drawing.Color]::Transparent
        $grpFixed.Controls.Add($lbl)

        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = '固定'
        $chk.Location = New-Object System.Drawing.Point(($offset + 40), 28)
        $chk.Size = New-Object System.Drawing.Size(52, 24)
        $chk.BackColor = [System.Drawing.Color]::Transparent
        $grpFixed.Controls.Add($chk)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Location = New-Object System.Drawing.Point(($offset + 98), 27)
        $box.Size = New-Object System.Drawing.Size($boxWidth, 26)
        $box.BorderStyle = 'FixedSingle'
        $box.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F1F5F9')
        $grpFixed.Controls.Add($box)

        # 勾选后才让输入框可输入并变白，视觉上直接区分「固定/未固定」
        # 注意：GetNewClosure 会新建作用域，闭包里不能用 $script: 取控件，必须直接捕获控件对象
        $chk.Add_CheckedChanged({
            if ($chk.Checked) {
                $box.BackColor = [System.Drawing.Color]::White
            } else {
                $box.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F1F5F9')
            }
            Apply-FixedToColumn -Index $slot
        }.GetNewClosure())
        $box.Add_TextChanged({ Apply-FixedToColumn -Index $slot }.GetNewClosure())

        $script:FixedChecks += $chk
        $script:FixedBoxes += $box
    }
}

function Apply-FixedToColumn {
    # 思路：勾选固定后把该值刷到所有行的这一列；不勾选则保留用户手改的内容
    param([int]$Index)
    $position = $Index - 1
    if ($position -lt 0 -or $position -ge $script:FixedChecks.Count) { return }
    if (-not $script:FixedChecks[$position].Checked) { return }
    if ($dgv.Columns.Count -le $Index) { return }
    $value = $script:FixedBoxes[$position].Text
    $script:SuppressRefresh = $true
    try {
        foreach ($gridRow in $dgv.Rows) {
            # 列 0 是文件名，段i 在第 i 列（不是 i+1），写错会整体串到下一段
            $gridRow.Cells[$Index].Value = $value
        }
    } finally {
        $script:SuppressRefresh = $false
    }
    try { [void](Refresh-ComposerStatus) } catch { $lblStatus.Text = "刷新失败：$($_.Exception.Message)" }
}

function Rebuild-Grid {
    $script:SegmentCount = [int]$numCount.Value
    $script:SuppressRefresh = $true
    try {
        Rebuild-FixedControls
        $dgv.Columns.Clear()
        $dgv.Rows.Clear()

        $nameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $nameColumn.HeaderText = '文件名'
        $nameColumn.Width = 200
        $nameColumn.ReadOnly = $true
        $nameColumn.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F3F5F9')
        [void]$dgv.Columns.Add($nameColumn)

        # 表格总宽 890：固定列占 200/220/100，剩余宽度按段数均分，段数多时不出现横向滚动条
        # 预留 10px 给表格边框/滚动条，否则刚好等于宽度也会挤出横向滚动条
        $segmentWidth = [math]::Max(60, [math]::Floor((890 - 200 - 220 - 100 - 10) / $script:SegmentCount))
        for ($segmentIndex = 1; $segmentIndex -le $script:SegmentCount; $segmentIndex++) {
            $segmentColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $segmentColumn.HeaderText = "段$segmentIndex"
            $segmentColumn.Width = $segmentWidth
            # 可编辑列用浅暖底，和只读列区分
            $segmentColumn.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFBEA')
            [void]$dgv.Columns.Add($segmentColumn)
        }

        $newNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $newNameColumn.HeaderText = '新文件名'
        $newNameColumn.Width = 220
        $newNameColumn.ReadOnly = $true
        $newNameColumn.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F3F5F9')
        [void]$dgv.Columns.Add($newNameColumn)

        $statusColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $statusColumn.HeaderText = '状态'
        $statusColumn.Width = 100
        $statusColumn.ReadOnly = $true
        $statusColumn.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
        [void]$dgv.Columns.Add($statusColumn)

        foreach ($item in $script:FileItems) {
            $rowIndex = $dgv.Rows.Add()
            $row = $dgv.Rows[$rowIndex]
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

function Refresh-ComposerStatus {
    # 思路：把表格内容转成计划，回填「新文件名」「状态」，并统计各类数量
    if ($dgv.Rows.Count -eq 0) {
        $lblStatus.Text = '当前文件夹没有可处理的文件。'
        return @()
    }
    $script:SuppressRefresh = $true
    try {
        $rows = @(Get-RowRecords -Grid $dgv -SegmentCount $script:SegmentCount -FolderPath $script:FolderPath)
        $plan = @(Build-ComposerPlan -Rows $rows -Separator $txtSeparator.Text)

        for ($index = 0; $index -lt $plan.Count; $index++) {
            $row = $dgv.Rows[$index]
            $row.Cells[$script:SegmentCount + 1].Value = $plan[$index].NewName
            $row.Cells[$script:SegmentCount + 2].Value = $plan[$index].Status
            if ($script:BlockedStatuses -contains $plan[$index].Status) {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
            } elseif ($plan[$index].Status -eq '未变化') {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
            } else {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }

        $executable = @($plan | Where-Object Status -eq '可执行').Count
        $unchanged = @($plan | Where-Object Status -eq '未变化').Count
        $blocked = @($plan | Where-Object { $script:BlockedStatuses -contains $_.Status }).Count
        $lblStatus.Text = "共 $($plan.Count) 个文件：可执行 $executable 个，未变化 $unchanged 个，有问题 $blocked 个。"
        # 有问题时状态文字转红，正常情况下用深灰
        if ($blocked -gt 0) {
            $lblStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#B91C1C')
        } else {
            $lblStatus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#374151')
        }
        return $plan
    } finally {
        $script:SuppressRefresh = $false
    }
}

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($txtFolder.Text.Trim().Length -gt 0 -and (Test-Path -LiteralPath $txtFolder.Text.Trim())) {
        $dialog.SelectedPath = $txtFolder.Text.Trim()
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dialog.SelectedPath
        try {
            Load-FileItems
            Rebuild-Grid
            [void](Refresh-ComposerStatus)
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '读取失败',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})

$numCount.Add_ValueChanged({
    if ($script:FileItems.Count -gt 0) {
        Rebuild-Grid
        [void](Refresh-ComposerStatus)
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
            foreach ($gridRow in $dgv.Rows) {
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
        [void](Refresh-ComposerStatus)
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '拆分失败',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnClearSegments.Add_Click({
    $script:SuppressRefresh = $true
    try {
        foreach ($gridRow in $dgv.Rows) {
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
    [void](Refresh-ComposerStatus)
})

$dgv.Add_CellValueChanged({
    if ($script:SuppressRefresh) { return }
    try { [void](Refresh-ComposerStatus) } catch {
        $lblStatus.Text = "刷新失败：$($_.Exception.Message)"
    }
})

$dgv.Add_CurrentCellDirtyStateChanged({
    if ($dgv.IsCurrentCellDirty) { $dgv.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
})

$btnRun.Add_Click({
    try {
        $plan = @(Refresh-ComposerStatus)
        if ($plan.Count -eq 0) { throw '没有可处理的文件。' }
        $blocked = @($plan | Where-Object { $script:BlockedStatuses -contains $_.Status })
        if ($blocked.Count -gt 0) {
            throw "有 $($blocked.Count) 行存在问题（红色行），请修正后再执行。"
        }
        $executable = @($plan | Where-Object Status -eq '可执行')
        if ($executable.Count -eq 0) { throw '没有需要改名的文件（全部未变化）。' }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "即将改名 $($executable.Count) 个文件，是否继续？`r`n改名后可在主界面点「撤销」回退。",
            '确认执行', [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $result = Invoke-ComposerRename -Plan $executable
        $message = "已改名 $($result.SuccessCount) 个，失败 $($result.Errors.Count) 个。"
        if ($result.Errors.Count -gt 0) {
            $message += "`r`n`r`n前 10 条错误：`r`n" + (($result.Errors | Select-Object -First 10) -join "`r`n")
        }
        [void][System.Windows.Forms.MessageBox]::Show($message, '执行结果',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        Load-FileItems
        Rebuild-Grid
        [void](Refresh-ComposerStatus)
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '无法执行',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$btnClose.Add_Click({ $form.Close() })

Initialize-ButtonStyle -Button $btnRun -Primary
Initialize-ButtonStyle -Button $btnClose

$form.Add_Shown({
    try {
        Load-FileItems
        Rebuild-Grid
        [void](Refresh-ComposerStatus)
    } catch {
        $lblStatus.Text = $_.Exception.Message
    }
    try { $form.Activate() } catch { }
})

if ($Diagnose) {
    # 界面已构建但不显示，直接验证「读取文件 → 建表 → 计算新名称」链路
    try {
        Load-FileItems
        Rebuild-Grid
        $plan = @(Refresh-ComposerStatus)
        Write-Output ("DIAG_ROWS={0} COLS={1} PLAN={2}" -f $dgv.Rows.Count, $dgv.Columns.Count, $plan.Count)
        foreach ($gridRow in $dgv.Rows) {
            $cells = @()
            foreach ($cell in $gridRow.Cells) { $cells += [string]$cell.Value }
            Write-Output ("  " + ($cells -join ' | '))
        }

        # 回归检查：段1 的固定值必须落在段1 列，串到段2 就说明列索引错位
        if ($dgv.Rows.Count -gt 0 -and $script:FixedChecks.Count -ge 2) {
            $script:FixedBoxes[0].Text = '固定甲'
            $script:FixedChecks[0].Checked = $true
            $landedValue = [string]$dgv.Rows[0].Cells[1].Value
            Write-Output ("FIXED_CHECK_CELL1={0}" -f $landedValue)
            $script:FixedChecks[0].Checked = $false
            $script:FixedBoxes[0].Text = ''
            if ($landedValue -ne '固定甲') { throw "固定值列错位：段1 列实际为 $landedValue" }
        }
    } finally {
        $form.Dispose()
    }
    return
}

try {
    [void]$form.ShowDialog()
} catch {
    $logPath = Join-Path $env:TEMP 'BatchRenamerComposer_error.log'
    "时间：$(Get-Date -Format s)`r`n错误：$($_.Exception.Message)`r`n堆栈：$($_.ScriptStackTrace)" |
        Set-Content -LiteralPath $logPath -Encoding UTF8
    [void][System.Windows.Forms.MessageBox]::Show("程序启动失败：$($_.Exception.Message)`r`n日志：$logPath", '错误',
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
