# -*- coding: utf-8 -*-
# ============================================================
#  表格比对工具 v5.0.0
#  功能: 逐列比对两个 Excel 文件，差异标红，MAC/IP自动归一化
# ============================================================

$scriptDir = $PSScriptRoot
$logDir = Join-Path $scriptDir "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$errLog = Join-Path $logDir ("error_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

try {

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$logFile = Join-Path $logDir ("compare_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

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
    $itemH = 20
    $maxVisible = 20
    $visible = [Math]::Min($items.Count, $maxVisible)
    $lstH = $visible * $itemH
    $extraH = if ($multiSelect) { 50 } else { 10 }
    $formH = $lstH + $extraH + 70
    $formW = 660

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size($formW, $formH)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = if ($multiSelect) { "提示: 按住 Ctrl 可多选，共 $($items.Count) 项" } else { "提示: 共 $($items.Count) 项，请选择一项" }
    $lblInfo.Location = New-Object System.Drawing.Point(16, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(620, 22)
    $lblInfo.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblInfo)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(16, 38)
    $lst.Size = New-Object System.Drawing.Size(612, $lstH)
    $lst.IntegralHeight = $false
    $lst.BorderStyle = "FixedSingle"
    if ($multiSelect) { $lst.SelectionMode = "MultiExtended" } else { $lst.SelectionMode = "One" }
    foreach ($item in $items) { $lst.Items.Add($item) | Out-Null }
    if (-not $multiSelect -and $items.Count -gt 0) { $lst.SelectedIndex = 0 }
    $form.Controls.Add($lst)

    $btnY = $lstH + 46

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

    $form.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
    })
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

# ============================================================
#  手动列配对
# ============================================================

function Show-ColumnMapper($origHeaders, $copyHeaders, $origSheetName, $copySheetName) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "手动列配对 - $origSheetName vs $copySheetName"
    $form.Size = New-Object System.Drawing.Size(780, 520)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

    $lblLeft = New-Object System.Windows.Forms.Label
    $lblLeft.Text = "原始文件列"
    $lblLeft.Location = New-Object System.Drawing.Point(16, 12)
    $lblLeft.Size = New-Object System.Drawing.Size(300, 20)
    $lblLeft.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblLeft)

    $lstLeft = New-Object System.Windows.Forms.ListBox
    $lstLeft.Location = New-Object System.Drawing.Point(16, 34)
    $lstLeft.Size = New-Object System.Drawing.Size(300, 200)
    $lstLeft.IntegralHeight = $false
    $lstLeft.BorderStyle = "FixedSingle"
    foreach ($h in $origHeaders) { $lstLeft.Items.Add($h.name) | Out-Null }
    $form.Controls.Add($lstLeft)

    $lblRight = New-Object System.Windows.Forms.Label
    $lblRight.Text = "对比文件列"
    $lblRight.Location = New-Object System.Drawing.Point(456, 12)
    $lblRight.Size = New-Object System.Drawing.Size(300, 20)
    $lblRight.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblRight)

    $lstRight = New-Object System.Windows.Forms.ListBox
    $lstRight.Location = New-Object System.Drawing.Point(456, 34)
    $lstRight.Size = New-Object System.Drawing.Size(300, 200)
    $lstRight.IntegralHeight = $false
    $lstRight.BorderStyle = "FixedSingle"
    foreach ($h in $copyHeaders) { $lstRight.Items.Add($h.name) | Out-Null }
    $form.Controls.Add($lstRight)

    $btnPair = New-Object System.Windows.Forms.Button
    $btnPair.Text = "配对 >>"
    $btnPair.Location = New-Object System.Drawing.Point(330, 100)
    $btnPair.Size = New-Object System.Drawing.Size(100, 32)
    $btnPair.FlatStyle = "System"
    $form.Controls.Add($btnPair)

    $lblPaired = New-Object System.Windows.Forms.Label
    $lblPaired.Text = "已配对的列："
    $lblPaired.Location = New-Object System.Drawing.Point(16, 248)
    $lblPaired.Size = New-Object System.Drawing.Size(200, 20)
    $lblPaired.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblPaired)

    $lstPaired = New-Object System.Windows.Forms.ListBox
    $lstPaired.Location = New-Object System.Drawing.Point(16, 270)
    $lstPaired.Size = New-Object System.Drawing.Size(580, 130)
    $lstPaired.IntegralHeight = $false
    $lstPaired.BorderStyle = "FixedSingle"
    $form.Controls.Add($lstPaired)

    $pairedList = [System.Collections.ArrayList]::new()

    $btnPair.Add_Click({
        if ($lstLeft.SelectedIndex -lt 0 -or $lstRight.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("请先在左右两侧各选一个列", "提示")
            return
        }
        $oIdx = $lstLeft.SelectedIndex
        $cIdx = $lstRight.SelectedIndex
        $oName = $lstLeft.Items[$oIdx]
        $cName = $lstRight.Items[$cIdx]
        foreach ($p in $pairedList) {
            if ($p.origIdx -eq $oIdx -or $p.copyIdx -eq $cIdx) {
                [System.Windows.Forms.MessageBox]::Show("该列已配对过", "提示")
                return
            }
        }
        $pairedList.Add(@{ origIdx = $oIdx; copyIdx = $cIdx; origName = $oName; copyName = $cName }) | Out-Null
        $lstPaired.Items.Add("$oName  <-->  $cName") | Out-Null
    })

    $btnUnpair = New-Object System.Windows.Forms.Button
    $btnUnpair.Text = "移除"
    $btnUnpair.Location = New-Object System.Drawing.Point(610, 270)
    $btnUnpair.Size = New-Object System.Drawing.Size(80, 32)
    $btnUnpair.FlatStyle = "System"
    $btnUnpair.Add_Click({
        if ($lstPaired.SelectedIndex -ge 0) {
            $idx = $lstPaired.SelectedIndex
            $lstPaired.Items.RemoveAt($idx)
            $pairedList.RemoveAt($idx)
        }
    })
    $form.Controls.Add($btnUnpair)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "确定"
    $btnOK.Location = New-Object System.Drawing.Point(280, 418)
    $btnOK.Size = New-Object System.Drawing.Size(100, 36)
    $btnOK.FlatStyle = "System"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)
    $form.AcceptButton = $btnOK

    $form.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
    })
    $form.Add_FormClosing({
        if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
    })

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $pairs = @()
    foreach ($p in $pairedList) {
        $pairs += @{
            origIdx  = $origHeaders[$p.origIdx].index
            copyIdx  = $copyHeaders[$p.copyIdx].index
            origName = $origHeaders[$p.origIdx].name
            copyName = $copyHeaders[$p.copyIdx].name
        }
    }
    return $pairs
}

# ============================================================
#  Sheet 手动配对
# ============================================================

function Show-SheetMapper($sheetNamesO, $sheetNamesC, $preFilled) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "手动配对工作表"
    $form.Size = New-Object System.Drawing.Size(580, 480)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

    $lblL = New-Object System.Windows.Forms.Label
    $lblL.Text = "原始文件 Sheet ($($sheetNamesO.Count)个)"
    $lblL.Location = New-Object System.Drawing.Point(16, 12)
    $lblL.Size = New-Object System.Drawing.Size(220, 20)
    $lblL.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblL)

    $lstL = New-Object System.Windows.Forms.ListBox
    $lstL.Location = New-Object System.Drawing.Point(16, 34)
    $lstL.Size = New-Object System.Drawing.Size(220, 160)
    $lstL.IntegralHeight = $false
    $lstL.BorderStyle = "FixedSingle"
    foreach ($n in $sheetNamesO) { $lstL.Items.Add($n) | Out-Null }
    $form.Controls.Add($lstL)

    $lblR = New-Object System.Windows.Forms.Label
    $lblR.Text = "对比文件 Sheet ($($sheetNamesC.Count)个)"
    $lblR.Location = New-Object System.Drawing.Point(336, 12)
    $lblR.Size = New-Object System.Drawing.Size(220, 20)
    $lblR.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblR)

    $lstR = New-Object System.Windows.Forms.ListBox
    $lstR.Location = New-Object System.Drawing.Point(336, 34)
    $lstR.Size = New-Object System.Drawing.Size(220, 160)
    $lstR.IntegralHeight = $false
    $lstR.BorderStyle = "FixedSingle"
    foreach ($n in $sheetNamesC) { $lstR.Items.Add($n) | Out-Null }
    $form.Controls.Add($lstR)

    $btnPair = New-Object System.Windows.Forms.Button
    $btnPair.Text = "配对 >>"
    $btnPair.Location = New-Object System.Drawing.Point(246, 90)
    $btnPair.Size = New-Object System.Drawing.Size(80, 30)
    $btnPair.FlatStyle = "System"
    $form.Controls.Add($btnPair)

    $lblPaired = New-Object System.Windows.Forms.Label
    $lblPaired.Text = "已配对："
    $lblPaired.Location = New-Object System.Drawing.Point(16, 206)
    $lblPaired.Size = New-Object System.Drawing.Size(200, 20)
    $lblPaired.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblPaired)

    $lstPaired = New-Object System.Windows.Forms.ListBox
    $lstPaired.Location = New-Object System.Drawing.Point(16, 228)
    $lstPaired.Size = New-Object System.Drawing.Size(440, 130)
    $lstPaired.IntegralHeight = $false
    $lstPaired.BorderStyle = "FixedSingle"
    $form.Controls.Add($lstPaired)

    $pairedList = [System.Collections.ArrayList]::new()

    # 预填已自动配对的
    if ($preFilled) {
        foreach ($pf in $preFilled) {
            $li = $sheetNamesO.IndexOf($pf.orig)
            $ri = $sheetNamesC.IndexOf($pf.copy)
            if ($li -ge 0 -and $ri -ge 0) {
                $pairedList.Add(@{ leftIdx = $li; rightIdx = $ri; leftName = $pf.orig; rightName = $pf.copy }) | Out-Null
                $lstPaired.Items.Add("$($pf.orig)  <-->  $($pf.copy)  [自动]") | Out-Null
            }
        }
    }

    $btnPair.Add_Click({
        if ($lstL.SelectedIndex -lt 0 -or $lstR.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("请先在左右两侧各选一个 Sheet", "提示")
            return
        }
        $li = $lstL.SelectedIndex
        $ri = $lstR.SelectedIndex
        $ln = $lstL.Items[$li]
        $rn = $lstR.Items[$ri]
        foreach ($p in $pairedList) {
            if ($p.leftIdx -eq $li -or $p.rightIdx -eq $ri) {
                [System.Windows.Forms.MessageBox]::Show("该 Sheet 已配对过", "提示")
                return
            }
        }
        $pairedList.Add(@{ leftIdx = $li; rightIdx = $ri; leftName = $ln; rightName = $rn }) | Out-Null
        $lstPaired.Items.Add("$ln  <-->  $rn") | Out-Null
    })

    $btnUnpair = New-Object System.Windows.Forms.Button
    $btnUnpair.Text = "移除"
    $btnUnpair.Location = New-Object System.Drawing.Point(470, 228)
    $btnUnpair.Size = New-Object System.Drawing.Size(80, 30)
    $btnUnpair.FlatStyle = "System"
    $btnUnpair.Add_Click({
        if ($lstPaired.SelectedIndex -ge 0) {
            $idx = $lstPaired.SelectedIndex
            $lstPaired.Items.RemoveAt($idx)
            $pairedList.RemoveAt($idx)
        }
    })
    $form.Controls.Add($btnUnpair)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "确定"
    $btnOK.Location = New-Object System.Drawing.Point(200, 375)
    $btnOK.Size = New-Object System.Drawing.Size(100, 36)
    $btnOK.FlatStyle = "System"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)
    $form.AcceptButton = $btnOK

    $form.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
    })
    $form.Add_FormClosing({
        if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
    })

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $pairs = @()
    foreach ($p in $pairedList) {
        $pairs += @{ orig = $sheetNamesO[$p.leftIdx]; copy = $sheetNamesC[$p.rightIdx] }
    }
    return $pairs
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

function Normalize-Mac($v) {
    $clean = $v -replace '[:\-\.\s]', ''
    $clean = $clean.ToUpper()
    if ($clean -match '^[0-9A-F]{12}$') {
        return ($clean -replace '(.{2})(.{2})(.{2})(.{2})(.{2})(.{2})', '$1:$2:$3:$4:$5:$6')
    }
    return $v
}

function Normalize-Ip($v) {
    $parts = $v.Split('.')
    if ($parts.Count -eq 4) {
        $normalized = @()
        foreach ($p in $parts) {
            $n = 0
            if ([int]::TryParse($p, [ref]$n)) {
                $normalized += $n.ToString()
            } else {
                return $v
            }
        }
        return ($normalized -join '.')
    }
    return $v
}

function Compare-Values($v1, $v2) {
    $e1 = if ($v1) { $v1.Trim() } else { "" }
    $e2 = if ($v2) { $v2.Trim() } else { "" }

    if ($e1 -eq "" -or $e2 -eq "") { return "skip" }

    $e1 = Normalize-Mac $e1
    $e2 = Normalize-Mac $e2

    $e1 = Normalize-Ip $e1
    $e2 = Normalize-Ip $e2

    $e1 = $e1.ToLower()
    $e2 = $e2.ToLower()

    if ($e1 -ne $e2) { return "diff" }
    return "same"
}

# ============================================================
#  主流程
# ============================================================

Write-Log "========== 开始 v5.0.0 =========="

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
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$copyPath = Join-Path $dir ("${name}_比对结果_${timestamp}${ext}")
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
$sheetPairs = @()

$selO = Show-ListPicker $sheetNamesO "原始文件 - 选择要比对的工作表（Ctrl多选 / 全选）" $true
if ($null -eq $selO) { Write-Log "取消"; exit }
$selectedO = @(); foreach ($idx in $selO) { $selectedO += $sheetNamesO[$idx] }

$selC = Show-ListPicker $sheetNamesC "对比文件 - 选择要比对的工作表（Ctrl多选 / 全选）" $true
if ($null -eq $selC) { Write-Log "取消"; exit }
$selectedC = @(); foreach ($idx in $selC) { $selectedC += $sheetNamesC[$idx] }

# 按名称自动配对
foreach ($so in $selectedO) {
    if ($selectedC -contains $so) {
        $sheetPairs += @{ orig = $so; copy = $so }
    }
}

Write-Log "自动配对: $($sheetPairs.Count) 个 Sheet"

# 弹出手动配对，列出全部 Sheet，已自动配对的预填进去
$sheetPairs = Show-SheetMapper $sheetNamesO $sheetNamesC $sheetPairs
if ($null -eq $sheetPairs -or $sheetPairs.Count -eq 0) {
    Write-Log "未配对任何 Sheet"
    [System.Windows.Forms.MessageBox]::Show("未配对任何 Sheet，无法比对", "提示")
    exit
}

Write-Log "最终配对: $($sheetPairs.Count) 个 Sheet"
foreach ($sp in $sheetPairs) { Write-Log "  $($sp.orig) <--> $($sp.copy)" }

# ============================================================
#  执行比对
# ============================================================

Write-Log "正在比对..."
$xl = New-ExcelEngine
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wbO = $xl.Workbooks.Open($fOriginal)
$wbC = $xl.Workbooks.Open($copyPath)

# 全局统计
$totalDiffCells = 0
$totalSkipCells = 0
$totalEmptyRows = 0
$totalAddedCount = 0
$totalMissingCount = 0
$allDiffSummary = @()

foreach ($pair in $sheetPairs) {
    $sheetO = $pair.orig
    $sheetC = $pair.copy
    Write-Log ""
    Write-Log "===== 正在比对 Sheet: $sheetO ====="

    $wsO = $wbO.Sheets($sheetO)
    $wsC = $wbC.Sheets($sheetC)

    $origMaxCol = 0; $origMaxRow = 0
    try { $origMaxCol = $wsO.UsedRange.Columns.Count; $origMaxRow = $wsO.UsedRange.Rows.Count } catch {}
    $origHeaderRow = Detect-HeaderRow $wsO $origMaxCol 10
    $origResult = Read-Headers $wsO $origHeaderRow $origMaxCol $origMaxRow
    $origHeaders = $origResult.headers
    $origRowCount = $origResult.rowCount
    Write-Log "原始表头行: $origHeaderRow，范围: $origMaxCol 列 x $origMaxRow 行，表头: $($origHeaders.Count) 个"

    $copyMaxCol = 0; $copyMaxRow = 0
    try { $copyMaxCol = $wsC.UsedRange.Columns.Count; $copyMaxRow = $wsC.UsedRange.Rows.Count } catch {}
    $copyHeaderRow = Detect-HeaderRow $wsC $copyMaxCol 10
    $copyResult = Read-Headers $wsC $copyHeaderRow $copyMaxCol $copyMaxRow
    $copyHeaders = $copyResult.headers
    $copyRowCount = $copyResult.rowCount
    Write-Log "对比表头行: $copyHeaderRow，范围: $copyMaxCol 列 x $copyMaxRow 行，表头: $($copyHeaders.Count) 个"

    if ($origHeaders.Count -eq 0 -or $copyHeaders.Count -eq 0) {
        Write-Log "警告: Sheet '$sheetO' 未找到表头，跳过"
        continue
    }

    # 选择列
    $origNames = @(); foreach ($h in $origHeaders) { $origNames += $h.name }
    $copyNames = @(); foreach ($h in $copyHeaders) { $copyNames += $h.name }

    $selO = Show-ListPicker $origNames "原始文件 [$sheetO] - 选择要比对的列（Ctrl多选）" $true
    if ($null -eq $selO) { Write-Log "取消"; try { $wbC.Close(0) } catch {}; try { $wbO.Close(0) } catch {}; Release-Excel $xl; return }
    $selC = Show-ListPicker $copyNames "对比文件 [$sheetC] - 选择要比对的列（Ctrl多选）" $true
    if ($null -eq $selC) { Write-Log "取消"; try { $wbC.Close(0) } catch {}; try { $wbO.Close(0) } catch {}; Release-Excel $xl; return }

    $origSelected = $selO
    $copySelected = $selC
    Write-Log "原始比对列: $($origSelected.Count)，对比比对列: $($copySelected.Count)"

    # 按名称自动配对
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
    Write-Log "自动配对: $($pairs.Count) 列"
    foreach ($p in $pairs) { Write-Log "  $($p.origName) <--> $($p.copyName)" }

    # 自动配对失败时，弹出手动配对
    if ($pairs.Count -eq 0) {
        Write-Log "表头名称不同，进入手动配对..."
        $pairs = Show-ColumnMapper $origHeaders $copyHeaders $sheetO $sheetC
        if ($null -eq $pairs -or $pairs.Count -eq 0) {
            Write-Log "手动配对为空或取消，跳过 Sheet '$sheetO'"
            continue
        }
        Write-Log "手动配对: $($pairs.Count) 列"
        foreach ($p in $pairs) { Write-Log "  $($p.origName) <--> $($p.copyName)" }
    }

    # ---- 执行比对 ----
    $diffCells = 0
    $skipCells = 0
    $emptyRows = 0
    $diffSummary = @()
    $startRow = [Math]::Max($origHeaderRow, $copyHeaderRow) + 1
    $maxRow = [Math]::Max($origRowCount, $copyRowCount)

    for ($r = $startRow; $r -le $maxRow; $r++) {
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
            $diffSummary += @{ type = "缺少"; row = $r; key = $r; col = "-"; old = "-"; new = "-"; sheet = $sheetO }
            continue
        }
        if (-not $hasOrig) {
            $diffSummary += @{ type = "新增"; row = $r; key = $r; col = "-"; old = "-"; new = "-"; sheet = $sheetO }
            continue
        }

        foreach ($p in $pairs) {
            $v1 = $wsO.Cells($r, $p.origIdx).Text
            $v2 = $wsC.Cells($r, $p.copyIdx).Text
            $result = Compare-Values $v1 $v2
            if ($result -eq "skip") { $skipCells++; continue }
            if ($result -eq "diff") {
                $diffCells++
                $cell = $wsC.Cells($r, $p.copyIdx)
                $cell.Interior.Color = 255
                $cell.Font.Color = 16777215
                $cell.Font.Bold = $true
                $diffSummary += @{ type = "差异"; row = $r; key = $r; col = $p.origName; old = $v1; new = $v2; sheet = $sheetO }
            }
        }
    }

    # 对新增/缺少行整行标色
    foreach ($d in $diffSummary) {
        if ($d.type -eq "新增") {
            for ($c = 1; $c -le $copyMaxCol; $c++) {
                try {
                    $cell = $wsC.Cells($d.row, $c)
                    $cell.Interior.Color = 65280
                    $cell.Font.Color = 16777215
                } catch {}
            }
        }
        if ($d.type -eq "缺少") {
            for ($c = 1; $c -le $copyMaxCol; $c++) {
                try {
                    $cell = $wsC.Cells($d.row, $c)
                    $cell.Interior.Color = 65535
                    $cell.Font.Color = 0
                } catch {}
            }
        }
    }

    # 在副本表末尾加「说明」列
    $descCol = $copyMaxCol + 1
    $wsC.Cells($copyHeaderRow, $descCol).Value = "说明"
    $wsC.Cells($copyHeaderRow, $descCol).Font.Bold = $true

    # 按行汇总差异说明
    $rowDescMap = @{}
    foreach ($d in $diffSummary) {
        $r = $d.row
        if (-not $rowDescMap.ContainsKey($r)) { $rowDescMap[$r] = @() }
        $rowDescMap[$r] += $d
    }
    foreach ($r in $rowDescMap.Keys) {
        $items = $rowDescMap[$r]
        $types = @(); foreach ($item in $items) { $types += $item.type }
        $hasDiff = $types -contains "差异"
        $hasNew = $types -contains "新增"
        $hasMiss = $types -contains "缺少"
        $desc = ""
        if ($hasDiff) {
            $diffCols = @(); foreach ($item in $items) { if ($item.type -eq "差异") { $diffCols += $item.col } }
            $desc = "值不同: $($diffCols -join ', ')"
        }
        if ($hasNew) { $desc = "对比文件多出的行" }
        if ($hasMiss) { $desc = "原始文件多出的行" }
        $wsC.Cells($r, $descCol).Value = $desc
    }

    $addedCount = ($diffSummary | Where-Object { $_.type -eq "新增" }).Count
    $missingCount = ($diffSummary | Where-Object { $_.type -eq "缺少" }).Count

    Write-Log "Sheet '$sheetO' 比对完成: 差异=$diffCells, 跳过空值=$skipCells, 跳过空行=$emptyRows, 新增=$addedCount, 缺少=$missingCount"

    $totalDiffCells += $diffCells
    $totalSkipCells += $skipCells
    $totalEmptyRows += $emptyRows
    $totalAddedCount += $addedCount
    $totalMissingCount += $missingCount

    $allDiffSummary += $diffSummary
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

Write-Host ""
Write-Host "========== 比对完成 =========="
Write-Host "  比对 Sheet 数: $($sheetPairs.Count)"
Write-Host "  差异单元格: $totalDiffCells（已标红）"
Write-Host "  跳过空值: $totalSkipCells"
Write-Host "  跳过空行: $totalEmptyRows"
Write-Host "  新增多出: $totalAddedCount 行（标绿）"
Write-Host "  缺少: $totalMissingCount 行（标黄）"
Write-Host "  结果文件: $copyPath"
Write-Host "  详细日志: $logFile"
Write-Host "=============================="

Write-Log "========== 结果 =========="
Write-Log "比对 Sheet 数: $($sheetPairs.Count)"
Write-Log "差异单元格: $totalDiffCells"
Write-Log "跳过空值: $totalSkipCells"
Write-Log "跳过空行: $totalEmptyRows"
Write-Log "新增行: $totalAddedCount"
Write-Log "缺少行: $totalMissingCount"
Write-Log "状态: 完成"
Write-Log "========== 结束 =========="

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "打开结果文件"
$btnOpen.Width = 140
$btnOpen.Height = 35

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "关闭"
$btnClose.Width = 100
$btnClose.Height = 35

$dlgResult = New-Object System.Windows.Forms.Form
$dlgResult.Text = "比对完成"
$dlgResult.Size = New-Object System.Drawing.Size(400, 180)
$dlgResult.StartPosition = "CenterScreen"
$dlgResult.FormBorderStyle = "FixedDialog"
$dlgResult.MaximizeBox = $false
$dlgResult.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "比对完成!`n`n" +
    "比对 Sheet 数: $($sheetPairs.Count)`n" +
    "差异单元格: $totalDiffCells`n" +
    "新增行: $totalAddedCount`n" +
    "缺少行: $totalMissingCount`n" +
    "结果文件已标记差异并添加「说明」列"
$lblResult.Location = New-Object System.Drawing.Point(16, 12)
$lblResult.Size = New-Object System.Drawing.Size(360, 90)
$dlgResult.Controls.Add($lblResult)

$btnOpen.Location = New-Object System.Drawing.Point(60, 110)
$btnClose.Location = New-Object System.Drawing.Point(230, 110)
$dlgResult.Controls.Add($btnOpen)
$dlgResult.Controls.Add($btnClose)

$btnOpen.Add_Click({ Start-Process $copyPath; $dlgResult.Close() })
$btnClose.Add_Click({ $dlgResult.Close() })

$dlgResult.ShowDialog() | Out-Null

} catch {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errMsg = "[$ts] $($_.Exception.Message)`n位置: 行$($_.InvocationInfo.ScriptLineNumber)`n$($_.ScriptStackTrace)"
    try { Add-Content -Path $errLog -Value $errMsg -Encoding UTF8 } catch {}
    try {
        [System.Windows.Forms.MessageBox]::Show("程序出错，日志已保存到:`n$errLog`n`n$($_.Exception.Message)", "错误")
    } catch {
        Write-Host $errMsg
        Write-Host "`n按任意键退出..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
