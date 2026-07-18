# -*- coding: utf-8 -*-
# ============================================================
#  表格比对工具 v2.0.0
#  功能: 逐列比对两个 Excel 文件，差异标红，支持关键列匹配
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
#  常量与日志
# ============================================================

$scriptDir = $PSScriptRoot
$logDir = Join-Path $scriptDir "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logFile = Join-Path $logDir ("compare_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
$configFile = Join-Path $scriptDir "config.json"

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

# ============================================================
#  工具函数
# ============================================================

function New-ExcelEngine {
    try { $xl = New-Object -ComObject Excel.Application; Write-Log "引擎: Excel"; return $xl } catch {}
    try { $xl = New-Object -ComObject Ket.Application; Write-Log "引擎: WPS(Ket)"; return $xl } catch {}
    try { $xl = New-Object -ComObject KWps.Application; Write-Log "引擎: WPS(KWps)"; return $xl } catch {}
    Write-Log "错误: 未检测到 Excel 或 WPS"
    [System.Windows.Forms.MessageBox]::Show("未检测到 Excel 或 WPS", "错误")
    exit 1
}

function Release-Excel($xl) {
    try { $xl.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Pick-File($title, $filter) {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $title
    $dlg.Filter = $filter
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.FileName }
    return $null
}

function Show-ListPicker($items, $title, $multiSelect) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(460, 380)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = if ($multiSelect) { "提示: 按住 Ctrl 可多选，共 $($items.Count) 项" } else { "提示: 共 $($items.Count) 项，请选择一项" }
    $lblInfo.Location = New-Object System.Drawing.Point(16, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(420, 22)
    $lblInfo.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblInfo)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(16, 38)
    $lst.Size = New-Object System.Drawing.Size(412, 230)
    $lst.IntegralHeight = $false
    $lst.BorderStyle = "FixedSingle"
    if ($multiSelect) { $lst.SelectionMode = "MultiExtended" } else { $lst.SelectionMode = "One" }
    foreach ($item in $items) { $lst.Items.Add($item) | Out-Null }
    if (-not $multiSelect -and $items.Count -gt 0) { $lst.SelectedIndex = 0 }
    $form.Controls.Add($lst)

    $btnY = 282

    if ($multiSelect) {
        $btnAll = New-Object System.Windows.Forms.Button
        $btnAll.Text = "全选"
        $btnAll.Location = New-Object System.Drawing.Point(16, $btnY)
        $btnAll.Size = New-Object System.Drawing.Size(80, 32)
        $btnAll.FlatStyle = "System"
        $btnAll.Add_Click({ for ($i = 0; $i -lt $lst.Items.Count; $i++) { $lst.SetSelected($i, $true) } })
        $form.Controls.Add($btnAll)

        $btnNone = New-Object System.Windows.Forms.Button
        $btnNone.Text = "全不选"
        $btnNone.Location = New-Object System.Drawing.Point(104, $btnY)
        $btnNone.Size = New-Object System.Drawing.Size(80, 32)
        $btnNone.FlatStyle = "System"
        $btnNone.Add_Click({ for ($i = 0; $i -lt $lst.Items.Count; $i++) { $lst.SetSelected($i, $false) } })
        $form.Controls.Add($btnNone)
    }

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "确定"
    $btn.Location = New-Object System.Drawing.Point(348, $btnY)
    $btn.Size = New-Object System.Drawing.Size(80, 32)
    $btn.FlatStyle = "System"
    $btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btn)
    $form.AcceptButton = $btn

    $form.CancelButton = $btn
    $form.Add_FormClosing({
        if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
    })

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    if ($multiSelect) {
        $selected = @()
        foreach ($idx in $lst.SelectedIndices) { $selected += $idx }
        return $selected
    } else {
        return @($lst.SelectedIndex)
    }
}

function Show-SettingsForm($savedConfig) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "表格比对工具 — 设置"
    $form.Size = New-Object System.Drawing.Size(420, 420)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

    $titleLbl = New-Object System.Windows.Forms.Label
    $titleLbl.Text = "比对选项设置"
    $titleLbl.Location = New-Object System.Drawing.Point(16, 12)
    $titleLbl.Size = New-Object System.Drawing.Size(380, 24)
    $titleLbl.Font = New-Object System.Drawing.Font("Microsoft YaHei", 11, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLbl)

    $sepLine = New-Object System.Windows.Forms.Label
    $sepLine.BorderStyle = "Fixed3D"
    $sepLine.Size = New-Object System.Drawing.Size(374, 2)
    $sepLine.Location = New-Object System.Drawing.Point(16, 38)
    $form.Controls.Add($sepLine)

    $y = 50

    $chkCase = New-Object System.Windows.Forms.CheckBox
    $chkCase.Text = "忽略大小写 (A 与 a 视为相同)"
    $chkCase.Location = New-Object System.Drawing.Point(24, $y)
    $chkCase.Size = New-Object System.Drawing.Size(370, 22)
    $chkCase.Checked = $savedConfig.ignoreCase
    $form.Controls.Add($chkCase)
    $y += 29

    $chkDate = New-Object System.Windows.Forms.CheckBox
    $chkDate.Text = "忽略日期格式差异 (2020/1/1 vs 2020-01-01)"
    $chkDate.Location = New-Object System.Drawing.Point(24, $y)
    $chkDate.Size = New-Object System.Drawing.Size(370, 22)
    $chkDate.Checked = $savedConfig.ignoreDateFormat
    $form.Controls.Add($chkDate)
    $y += 29

    $lblTol = New-Object System.Windows.Forms.Label
    $lblTol.Text = "数值容差:"
    $lblTol.Location = New-Object System.Drawing.Point(44, $y)
    $lblTol.Size = New-Object System.Drawing.Size(68, 24)
    $lblTol.TextAlign = "MiddleLeft"
    $form.Controls.Add($lblTol)

    $txtTol = New-Object System.Windows.Forms.TextBox
    $txtTol.Text = [string]$savedConfig.tolerance
    $txtTol.Location = New-Object System.Drawing.Point(112, $y)
    $txtTol.Size = New-Object System.Drawing.Size(80, 24)
    $form.Controls.Add($txtTol)

    $lblTolHint = New-Object System.Windows.Forms.Label
    $lblTolHint.Text = "0=精确比对，如填 0.01 表示差在该值以内视为相同"
    $lblTolHint.Location = New-Object System.Drawing.Point(198, $y)
    $lblTolHint.Size = New-Object System.Drawing.Size(196, 24)
    $lblTolHint.TextAlign = "MiddleLeft"
    $lblTolHint.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblTolHint)
    $y += 33

    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.BorderStyle = "Fixed3D"
    $sep2.Size = New-Object System.Drawing.Size(374, 2)
    $sep2.Location = New-Object System.Drawing.Point(16, $y)
    $form.Controls.Add($sep2)
    $y += 14
    $chkExclude = New-Object System.Windows.Forms.CheckBox
    $chkExclude.Text = "排除列模式 (选择无需比对的列，其余自动比对)"
    $chkExclude.Location = New-Object System.Drawing.Point(24, $y)
    $chkExclude.Size = New-Object System.Drawing.Size(370, 22)
    $chkExclude.Checked = $savedConfig.excludeMode
    $form.Controls.Add($chkExclude)
    $y += 29

    $chkKeyCol = New-Object System.Windows.Forms.CheckBox
    $chkKeyCol.Text = "关键列匹配行 (按指定列值匹配，不依赖行号顺序)"
    $chkKeyCol.Location = New-Object System.Drawing.Point(24, $y)
    $chkKeyCol.Size = New-Object System.Drawing.Size(370, 22)
    $chkKeyCol.Checked = $savedConfig.useKeyColumn
    $form.Controls.Add($chkKeyCol)
    $y += 29

    $chkSummary = New-Object System.Windows.Forms.CheckBox
    $chkSummary.Text = "生成差异汇总 Sheet (在结果文件中追加一个差异明细表)"
    $chkSummary.Location = New-Object System.Drawing.Point(24, $y)
    $chkSummary.Size = New-Object System.Drawing.Size(370, 22)
    $chkSummary.Checked = $savedConfig.createSummary
    $form.Controls.Add($chkSummary)
    $y += 33

    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.BorderStyle = "Fixed3D"
    $sep3.Size = New-Object System.Drawing.Size(374, 2)
    $sep3.Location = New-Object System.Drawing.Point(16, $y)
    $form.Controls.Add($sep3)
    $y += 14

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "表头行号:"
    $lblHeader.Location = New-Object System.Drawing.Point(24, $y)
    $lblHeader.Size = New-Object System.Drawing.Size(72, 24)
    $lblHeader.TextAlign = "MiddleLeft"
    $form.Controls.Add($lblHeader)

    $txtHeader = New-Object System.Windows.Forms.TextBox
    $txtHeader.Text = [string]$savedConfig.headerRow
    $txtHeader.Location = New-Object System.Drawing.Point(100, $y)
    $txtHeader.Size = New-Object System.Drawing.Size(80, 24)
    $form.Controls.Add($txtHeader)

    $lblHdrHint = New-Object System.Windows.Forms.Label
    $lblHdrHint.Text = "0=自动检测，其他数字=指定行号"
    $lblHdrHint.Location = New-Object System.Drawing.Point(186, $y)
    $lblHdrHint.Size = New-Object System.Drawing.Size(200, 24)
    $lblHdrHint.TextAlign = "MiddleLeft"
    $lblHdrHint.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblHdrHint)
    $y += 36

    $sep4 = New-Object System.Windows.Forms.Label
    $sep4.BorderStyle = "Fixed3D"
    $sep4.Size = New-Object System.Drawing.Size(374, 2)
    $sep4.Location = New-Object System.Drawing.Point(16, $y)
    $form.Controls.Add($sep4)
    $y += 16

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "开始比对"
    $btn.Location = New-Object System.Drawing.Point(286, $y)
    $btn.Size = New-Object System.Drawing.Size(104, 34)
    $btn.FlatStyle = "System"
    $btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btn)
    $form.AcceptButton = $btn

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消"
    $btnCancel.Location = New-Object System.Drawing.Point(174, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(104, 34)
    $btnCancel.FlatStyle = "System"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $form.Tag = @{ chkCase = $chkCase; txtTol = $txtTol; chkDate = $chkDate; chkExclude = $chkExclude; chkKeyCol = $chkKeyCol; chkSummary = $chkSummary; txtHeader = $txtHeader }

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $tol = 0
    [double]::TryParse($txtTol.Text, [ref]$tol) | Out-Null
    $hdr = 0
    [int]::TryParse($txtHeader.Text, [ref]$hdr) | Out-Null

    return @{
        ignoreCase = $chkCase.Checked
        tolerance = $tol
        ignoreDateFormat = $chkDate.Checked
        excludeMode = $chkExclude.Checked
        useKeyColumn = $chkKeyCol.Checked
        createSummary = $chkSummary.Checked
        headerRow = $hdr
    }
}

function Load-Config {
    $default = @{
        ignoreCase = $false
        tolerance = 0
        ignoreDateFormat = $false
        excludeMode = $false
        useKeyColumn = $false
        createSummary = $false
        headerRow = 0
    }
    if (Test-Path $configFile) {
        try {
            $json = Get-Content $configFile -Raw -Encoding UTF8
            $obj = $json | ConvertFrom-Json
            $default.ignoreCase = [bool]$obj.ignoreCase
            $default.tolerance = [double]$obj.tolerance
            $default.ignoreDateFormat = [bool]$obj.ignoreDateFormat
            $default.excludeMode = [bool]$obj.excludeMode
            $default.useKeyColumn = [bool]$obj.useKeyColumn
            $default.createSummary = if ($null -ne $obj.createSummary) { [bool]$obj.createSummary } else { $false }
            $default.headerRow = [int]$obj.headerRow
        } catch {}
    }
    return $default
}

function Save-Config($cfg) {
    try {
        $cfg | ConvertTo-Json | Set-Content $configFile -Encoding UTF8
    } catch {}
}

# ============================================================
#  表头检测
# ============================================================

function Detect-HeaderRow($ws, $maxCol, $maxScanRow) {
    $bestRow = 1
    $bestCount = 0
    for ($r = 1; $r -le $maxScanRow; $r++) {
        $count = 0
        for ($c = 1; $c -le $maxCol; $c++) {
            $v = ""
            try { $v = $ws.Cells($r, $c).Text } catch {}
            if ($v -and $v.Trim() -ne "") { $count++ }
        }
        if ($count -gt $bestCount) { $bestCount = $count; $bestRow = $r }
    }
    return $bestRow
}

function Read-Headers($ws, $headerRow, $maxCol, $maxRow) {
    $headers = @()
    for ($i = 1; $i -le $maxCol; $i++) {
        $h = ""
        try { $h = $ws.Cells($headerRow, $i).Text } catch {}
        if ($h -and $h.Trim() -ne "") { $headers += @{ index = $i; name = "Col$i - $h"; field = $h.Trim() } }
    }
    return @{ headers = $headers; rowCount = $maxRow; headerRow = $headerRow }
}

# ============================================================
#  比对逻辑
# ============================================================

function Compare-Values($v1, $v2, $cfg) {
    $e1 = if ($v1) { $v1.Trim() } else { "" }
    $e2 = if ($v2) { $v2.Trim() } else { "" }

    if ($e1 -eq "" -or $e2 -eq "") { return "skip" }

    if ($cfg.ignoreCase) {
        $e1 = $e1.ToLower()
        $e2 = $e2.ToLower()
    }

    if ($cfg.ignoreDateFormat) {
        $d1 = [datetime]::MinValue
        $d2 = [datetime]::MinValue
        $parsed1 = [datetime]::TryParse($e1, [ref]$d1)
        $parsed2 = [datetime]::TryParse($e2, [ref]$d2)
        if ($parsed1 -and $parsed2) {
            if ($d1 -eq $d2) { return "same" }
        }
    }

    if ($cfg.tolerance -gt 0) {
        $n1 = 0; $n2 = 0
        $parsed1 = [double]::TryParse($e1, [ref]$n1)
        $parsed2 = [double]::TryParse($e2, [ref]$n2)
        if ($parsed1 -and $parsed2) {
            if ([Math]::Abs($n1 - $n2) -le $cfg.tolerance) { return "same" }
        }
    }

    if ($e1 -ne $e2) { return "diff" }
    return "same"
}

# ============================================================
#  主流程
# ============================================================

Write-Log "========== 开始 v2.0.0 =========="

# 加载配置
$config = Load-Config

# 设置界面
$settings = Show-SettingsForm $config
if ($null -eq $settings) { Write-Log "取消: 设置"; exit }
$config = $settings
Save-Config $config
Write-Log "设置: 忽略大小写=$($config.ignoreCase) 容差=$($config.tolerance) 忽略日期=$($config.ignoreDateFormat) 排除模式=$($config.excludeMode) 关键列=$($config.useKeyColumn) 汇总=$($config.createSummary) 表头行=$($config.headerRow)"

# 选择文件
$filter = "Excel|*.xlsx;*.xls;*.xlsm|所有文件|*.*"
$fOriginal = Pick-File "选择原始文件" $filter
if (-not $fOriginal) { Write-Log "取消"; exit }
Write-Log "原始文件: $fOriginal"

$fCompare = Pick-File "选择对比文件" $filter
if (-not $fCompare) { Write-Log "取消"; exit }
Write-Log "对比文件: $fCompare"

$dir = [System.IO.Path]::GetDirectoryName($fCompare)
$name = [System.IO.Path]::GetFileNameWithoutExtension($fCompare)
$ext = [System.IO.Path]::GetExtension($fCompare)
$copyPath = Join-Path $dir ("${name}_marked${ext}")
Copy-Item $fCompare $copyPath -Force
Write-Log "已创建副本: $copyPath"

# 打开Excel读取工作表
$xl = New-ExcelEngine
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wbO = $xl.Workbooks.Open($fOriginal)
$sheetNamesO = @()
foreach ($s in $wbO.Sheets) { $sheetNamesO += $s.Name }
$wbO.Close(0)

$wbC = $xl.Workbooks.Open($copyPath)
$sheetNamesC = @()
foreach ($s in $wbC.Sheets) { $sheetNamesC += $s.Name }
$wbC.Close(0)
Release-Excel $xl

# 选择工作表
if ($sheetNamesO.Count -gt 1) {
    $sel = Show-ListPicker $sheetNamesO "选择原始文件的工作表" $false
    if ($null -eq $sel) { Write-Log "取消"; exit }
    $sheetO = $sheetNamesO[$sel[0]]
} else {
    $sheetO = $sheetNamesO[0]
    Write-Log "原始文件仅一个工作表，自动选择: $sheetO"
}

if ($sheetNamesC.Count -gt 1) {
    $sel = Show-ListPicker $sheetNamesC "选择对比文件的工作表" $false
    if ($null -eq $sel) { Write-Log "取消"; exit }
    $sheetC = $sheetNamesC[$sel[0]]
} else {
    $sheetC = $sheetNamesC[0]
    Write-Log "对比文件仅一个工作表，自动选择: $sheetC"
}

# 读取表头
Write-Log "正在读取表头..."
$xl = New-ExcelEngine
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wbO = $xl.Workbooks.Open($fOriginal)
$wsO = $wbO.Sheets($sheetO)
$origMaxCol = 0; $origMaxRow = 0
try { $origMaxCol = $wsO.UsedRange.Columns.Count; $origMaxRow = $wsO.UsedRange.Rows.Count } catch {}
$origHeaderRow = if ($config.headerRow -gt 0) { $config.headerRow } else { Detect-HeaderRow $wsO $origMaxCol 10 }
Write-Log "原始表头行: $origHeaderRow，范围: $origMaxCol 列 x $origMaxRow 行"
$origResult = Read-Headers $wsO $origHeaderRow $origMaxCol $origMaxRow
$origHeaders = $origResult.headers
$origRowCount = $origResult.rowCount
$wbO.Close(0)
Write-Log "原始表头: $($origHeaders.Count) 个"

$wbC = $xl.Workbooks.Open($copyPath)
$wsC = $wbC.Sheets($sheetC)
$copyMaxCol = 0; $copyMaxRow = 0
try { $copyMaxCol = $wsC.UsedRange.Columns.Count; $copyMaxRow = $wsC.UsedRange.Rows.Count } catch {}
$copyHeaderRow = if ($config.headerRow -gt 0) { $config.headerRow } else { Detect-HeaderRow $wsC $copyMaxCol 10 }
Write-Log "对比表头行: $copyHeaderRow，范围: $copyMaxCol 列 x $copyMaxRow 行"
$copyResult = Read-Headers $wsC $copyHeaderRow $copyMaxCol $copyMaxRow
$copyHeaders = $copyResult.headers
$copyRowCount = $copyResult.rowCount
$wbC.Close(0)
Write-Log "对比表头: $($copyHeaders.Count) 个"

Release-Excel $xl

if ($origHeaders.Count -eq 0 -or $copyHeaders.Count -eq 0) {
    Write-Log "错误: 未找到表头"
    [System.Windows.Forms.MessageBox]::Show("前10行中未找到表头", "错误")
    exit 1
}

# 选择列
$origNames = @(); foreach ($h in $origHeaders) { $origNames += $h.name }
$copyNames = @(); foreach ($h in $copyHeaders) { $copyNames += $h.name }

$modeLabel = if ($config.excludeMode) { "排除" } else { "比对" }
$selO = Show-ListPicker $origNames "原始文件 - 选择要${modeLabel}的列（Ctrl多选）" $true
if ($null -eq $selO) { Write-Log "取消"; exit }

$selC = Show-ListPicker $copyNames "对比文件 - 选择要${modeLabel}的列（Ctrl多选）" $true
if ($null -eq $selC) { Write-Log "取消"; exit }

# 处理排除模式
$origSelected = @()
$copySelected = @()
if ($config.excludeMode) {
    for ($i = 0; $i -lt $origHeaders.Count; $i++) {
        if ($selO -notcontains $i) { $origSelected += $i }
    }
    for ($i = 0; $i -lt $copyHeaders.Count; $i++) {
        if ($selC -notcontains $i) { $copySelected += $i }
    }
} else {
    $origSelected = $selO
    $copySelected = $selC
}
Write-Log "原始比对列: $($origSelected.Count)，对比比对列: $($copySelected.Count)"

# 按名称配对
$pairs = @()
foreach ($oi in $origSelected) {
    $oField = $origHeaders[$oi].field
    foreach ($ci in $copySelected) {
        if ($copyHeaders[$ci].field -eq $oField) {
            $pairs += @{ origIdx = $origHeaders[$oi].index; copyIdx = $copyHeaders[$ci].index; origName = $origHeaders[$oi].name; copyName = $copyHeaders[$ci].name }
            break
        }
    }
}
Write-Log "配对: $($pairs.Count) 列"
foreach ($p in $pairs) { Write-Log "  $($p.origName) <--> $($p.copyName)" }

if ($pairs.Count -eq 0) {
    Write-Log "错误: 未找到名称相同的列"
    [System.Windows.Forms.MessageBox]::Show("未找到名称相同的列", "错误")
    exit 1
}

# 关键列选择
$keyPair = $null
if ($config.useKeyColumn) {
    Write-Log "请选择关键列..."
    $keyNamesO = @(); foreach ($h in $origHeaders) { $keyNamesO += $h.name }
    $keyNamesC = @(); foreach ($h in $copyHeaders) { $keyNamesC += $h.name }
    $keySelO = Show-ListPicker $keyNamesO "选择原始文件的关键列（用于匹配行）" $false
    if ($null -eq $keySelO) { Write-Log "取消"; exit }
    $keySelC = Show-ListPicker $keyNamesC "选择对比文件的关键列（用于匹配行）" $false
    if ($null -eq $keySelC) { Write-Log "取消"; exit }
    $keyPair = @{
        origIdx = $origHeaders[$keySelO[0]].index
        copyIdx = $copyHeaders[$keySelC[0]].index
        origName = $origHeaders[$keySelO[0]].name
        copyName = $copyHeaders[$keySelC[0]].name
    }
    Write-Log "关键列: $($keyPair.origName) <--> $($keyPair.copyName)"
}

# ============================================================
#  执行比对
# ============================================================

Write-Log "正在比对..."
$xl = New-ExcelEngine
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wbO = $xl.Workbooks.Open($fOriginal)
$wsO = $wbO.Sheets($sheetO)
$wbC = $xl.Workbooks.Open($copyPath)
$wsC = $wbC.Sheets($sheetC)

$diffCells = 0
$skipCells = 0
$totalCells = 0
$emptyRows = 0
$startRow = [Math]::Max($origHeaderRow, $copyHeaderRow) + 1

# 差异汇总数据
$diffSummary = @()

if ($config.useKeyColumn -and $keyPair) {
    # ---- 关键列匹配模式 ----
    Write-Log "模式: 关键列匹配"

    $origKeyMap = @{}
    for ($r = $startRow; $r -le $origRowCount; $r++) {
        $kv = ""
        try { $kv = $wsO.Cells($r, $keyPair.origIdx).Text } catch {}
        if ($kv -and $kv.Trim() -ne "") { $origKeyMap[$kv.Trim()] = $r }
    }

    $copyKeyMap = @{}
    for ($r = $startRow; $r -le $copyRowCount; $r++) {
        $kv = ""
        try { $kv = $wsC.Cells($r, $keyPair.copyIdx).Text } catch {}
        if ($kv -and $kv.Trim() -ne "") { $copyKeyMap[$kv.Trim()] = $r }
    }

    # 检查对比文件每行
    for ($r = $startRow; $r -le $copyRowCount; $r++) {
        $ckv = ""
        try { $ckv = $wsC.Cells($r, $keyPair.copyIdx).Text } catch {}
        $ckv = if ($ckv) { $ckv.Trim() } else { "" }
        if ($ckv -eq "") { $emptyRows++; continue }

        if (-not $origKeyMap.ContainsKey($ckv)) {
            # 对比文件多出的行
            $diffSummary += @{ type = "新增"; row = $r; key = $ckv; col = "-"; old = "-"; new = "-" }
            continue
        }

        $origRow = $origKeyMap[$ckv]
        foreach ($p in $pairs) {
            if ($p.origIdx -eq $keyPair.origIdx) { continue }
            $v1 = $wsO.Cells($origRow, $p.origIdx).Text
            $v2 = $wsC.Cells($r, $p.copyIdx).Text
            $totalCells++
            $result = Compare-Values $v1 $v2 $config
            if ($result -eq "skip") { $skipCells++; continue }
            if ($result -eq "diff") {
                $diffCells++
                $cell = $wsC.Cells($r, $p.copyIdx)
                $cell.Interior.Color = 255
                $cell.Font.Color = 16777215
                $cell.Font.Bold = $true
                $diffSummary += @{ type = "差异"; row = $r; key = $ckv; col = $p.origName; old = $v1; new = $v2 }
            }
        }
    }

    # 检查原始文件多出的行
    foreach ($okv in $origKeyMap.Keys) {
        if (-not $copyKeyMap.ContainsKey($okv)) {
            $diffSummary += @{ type = "缺少"; row = $origKeyMap[$okv]; key = $okv; col = "-"; old = "-"; new = "-" }
        }
    }
} else {
    # ---- 行号顺序模式 ----
    Write-Log "模式: 行号顺序"
    $maxRow = [Math]::Max($origRowCount, $copyRowCount)

    for ($r = $startRow; $r -le $maxRow; $r++) {
        # 检查全空行
        $allEmpty = $true
        foreach ($p in $pairs) {
            $tv1 = ""; $tv2 = ""
            try { $tv1 = $wsO.Cells($r, $p.origIdx).Text } catch {}
            try { $tv2 = $wsC.Cells($r, $p.copyIdx).Text } catch {}
            if ($tv1 -and $tv1.Trim() -ne "") { $allEmpty = $false }
            if ($tv2 -and $tv2.Trim() -ne "") { $allEmpty = $false }
        }
        if ($allEmpty) { $emptyRows++; continue }

        $hasOrig = $r -le $origRowCount
        $hasCopy = $r -le $copyRowCount

        if (-not $hasCopy) {
            $diffSummary += @{ type = "缺少"; row = $r; key = $r; col = "-"; old = "-"; new = "-" }
            continue
        }
        if (-not $hasOrig) {
            $diffSummary += @{ type = "新增"; row = $r; key = $r; col = "-"; old = "-"; new = "-" }
            continue
        }

        foreach ($p in $pairs) {
            $v1 = $wsO.Cells($r, $p.origIdx).Text
            $v2 = $wsC.Cells($r, $p.copyIdx).Text
            $totalCells++
            $result = Compare-Values $v1 $v2 $config
            if ($result -eq "skip") { $skipCells++; continue }
            if ($result -eq "diff") {
                $diffCells++
                $cell = $wsC.Cells($r, $p.copyIdx)
                $cell.Interior.Color = 255
                $cell.Font.Color = 16777215
                $cell.Font.Bold = $true
                $diffSummary += @{ type = "差异"; row = $r; key = $r; col = $p.origName; old = $v1; new = $v2 }
            }
        }
    }
}

# ============================================================
#  差异汇总Sheet
# ============================================================

if ($config.createSummary -and $diffSummary.Count -gt 0) {
    Write-Log "正在生成差异汇总..."
    $wsSummary = $wbC.Sheets.Add()
    $wsSummary.Name = "差异汇总"

    $wsSummary.Cells(1, 1).Value = "类型"
    $wsSummary.Cells(1, 2).Value = "行号"
    $wsSummary.Cells(1, 3).Value = "关键值"
    $wsSummary.Cells(1, 4).Value = "列名"
    $wsSummary.Cells(1, 5).Value = "原值"
    $wsSummary.Cells(1, 6).Value = "对比值"

    for ($i = 0; $i -lt $diffSummary.Count; $i++) {
        $d = $diffSummary[$i]
        $row = $i + 2
        $wsSummary.Cells($row, 1).Value = $d.type
        $wsSummary.Cells($row, 2).Value = $d.row
        $wsSummary.Cells($row, 3).Value = $d.key
        $wsSummary.Cells($row, 4).Value = $d.col
        $wsSummary.Cells($row, 5).Value = $d.old
        $wsSummary.Cells($row, 6).Value = $d.new

        if ($d.type -eq "差异") {
            $wsSummary.Cells($row, 1).Interior.Color = 255
            $wsSummary.Cells($row, 1).Font.Color = 16777215
        } elseif ($d.type -eq "新增") {
            $wsSummary.Cells($row, 1).Interior.Color = 65280
            $wsSummary.Cells($row, 1).Font.Color = 16777215
        } elseif ($d.type -eq "缺少") {
            $wsSummary.Cells($row, 1).Interior.Color = 16776960
        }
    }

    $wsSummary.Columns(1).ColumnWidth = 10
    $wsSummary.Columns(2).ColumnWidth = 8
    $wsSummary.Columns(3).ColumnWidth = 15
    $wsSummary.Columns(4).ColumnWidth = 20
    $wsSummary.Columns(5).ColumnWidth = 25
    $wsSummary.Columns(6).ColumnWidth = 25
}

# 保存
try {
    $wbC.Save()
    Write-Log "保存成功: $copyPath"
} catch {
    Write-Log "错误: 保存失败 - $($_.Exception.Message)"
}

$wbC.Close(0)
$wbO.Close(0)
Release-Excel $xl

# ============================================================
#  输出结果
# ============================================================

$addedCount = ($diffSummary | Where-Object { $_.type -eq "新增" }).Count
$missingCount = ($diffSummary | Where-Object { $_.type -eq "缺少" }).Count

Write-Host ""
Write-Host "========== 比对完成 =========="
Write-Host "  配对列数: $($pairs.Count)"
Write-Host "  差异单元格: $diffCells（已标红）"
Write-Host "  跳过空值: $skipCells"
Write-Host "  跳过空行: $emptyRows"
Write-Host "  新增多出: $addedCount 行（标绿）"
Write-Host "  缺少: $missingCount 行（标黄）"
Write-Host "  结果文件: $copyPath"
Write-Host "  详细日志: $logFile"
Write-Host "=============================="

Write-Log "========== 结果 =========="
Write-Log "配对列: $($pairs.Count)"
Write-Log "差异单元格: $diffCells"
Write-Log "跳过空值: $skipCells"
Write-Log "跳过空行: $emptyRows"
Write-Log "新增行: $addedCount"
Write-Log "缺少行: $missingCount"
Write-Log "状态: 完成"
Write-Log "========== 结束 =========="

$resultDlg = [System.Windows.Forms.MessageBox]::Show(
    "比对完成!`n`n" +
    "配对列数: $($pairs.Count)`n" +
    "差异单元格: $diffCells (已标红)`n" +
    "新增行: $addedCount (标绿)`n" +
    "缺少行: $missingCount (标黄)`n" +
    "跳过空值: $skipCells`n" +
    "跳过空行: $emptyRows`n`n" +
    "是否打开结果文件?",
    "比对完成",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

if ($resultDlg -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $copyPath
}
