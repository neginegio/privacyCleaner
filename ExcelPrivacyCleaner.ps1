Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$ErrorActionPreference = "Stop"

$script:SelectedFile = $null
$script:TempDir = $null
$script:TempWorkbook = $null
$script:Findings = New-Object System.Collections.ArrayList
$script:AliasMaps = @{
    Name = @{}
    Address = @{}
    Text = @{}
}

function Get-SafeText {
    param($Value)
    if ($null -eq $Value) { return "" }
    return [string]$Value
}

function Get-NormalizedAliasKey {
    param(
        [ValidateSet("Name", "Address", "Text")]
        [string]$Type,
        [string]$Original
    )
    $key = $Original.Trim()
    if ($Type -eq "Name") {
        return ($key -replace "[\s　]+", "")
    }
    if ($Type -eq "Address") {
        return ($key -replace "[\s　]+", "" -replace "[‐-―−ー－]", "-")
    }
    return $key
}

function Clear-TempWorkbook {
    if ($script:TempDir -and (Test-Path -LiteralPath $script:TempDir)) {
        try {
            Remove-Item -LiteralPath $script:TempDir -Recurse -Force
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "一時ファイルの削除に失敗しました。`n$($_.Exception.Message)",
                "一時ファイル削除エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }
    $script:TempDir = $null
    $script:TempWorkbook = $null
}

function Reset-State {
    $script:Findings.Clear()
    $script:AliasMaps = @{
        Name = @{}
        Address = @{}
        Text = @{}
    }
}

function Get-Alias {
    param(
        [ValidateSet("Name", "Address", "Text")]
        [string]$Type,
        [string]$Original
    )
    $key = Get-NormalizedAliasKey -Type $Type -Original $Original
    if (-not $script:AliasMaps[$Type].ContainsKey($key)) {
        $count = $script:AliasMaps[$Type].Count + 1
        $prefix = switch ($Type) {
            "Name" { "仮名" }
            "Address" { "住所" }
            default { "伏字" }
        }
        $script:AliasMaps[$Type][$key] = "{0}{1:000}" -f $prefix, $count
    }
    return $script:AliasMaps[$Type][$key]
}

function Test-NameHeader {
    param([string]$Header)
    return $Header -match "(氏名|名前|お名前|顧客名|患者名|社員名|担当者|宛名|申請者|利用者|名義)"
}

function Test-AddressHeader {
    param([string]$Header)
    return $Header -match "(住所|所在地|居所|番地|都道府県|市区町村|町名|地番|宛先|連絡先)"
}

function Get-DetectionCandidates {
    param([string]$Text)

    $items = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Text)) { return $items }

    $prefecture = "(北海道|東京都|京都府|大阪府|(?:青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|神奈川|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|兵庫|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|熊本|大分|宮崎|鹿児島|沖縄)県)"
    $addressPatterns = @(
        "〒?\s*\d{3}-?\d{4}\s*${prefecture}?[一-龯ぁ-んァ-ヶー0-9０-９\-－丁目番地号の\s　、,\.．]+",
        "${prefecture}[一-龯ぁ-んァ-ヶー0-9０-９\-－丁目番地号の\s　、,\.．]+",
        "[一-龯ぁ-んァ-ヶー]+(?:市|区|町|村)[一-龯ぁ-んァ-ヶー0-9０-９\-－丁目番地号の\s　、,\.．]+"
    )
    foreach ($pattern in $addressPatterns) {
        foreach ($match in [regex]::Matches($Text, $pattern)) {
            $value = $match.Value.Trim()
            if ($value.Length -ge 5) {
                [void]$items.Add([pscustomobject]@{
                    Type = "Address"
                    Value = $value
                    Reason = "自由記述欄の住所らしき文字列"
                })
            }
        }
    }

    $namePatterns = @(
        "(?:氏名|名前|お名前|顧客名|患者名|社員名|担当者|宛名|申請者|利用者|名義)\s*[:：]\s*([一-龯々〆ヵヶぁ-んァ-ヶー]{1,8}(?:[ 　]+[一-龯々〆ヵヶぁ-んァ-ヶー]{1,8})?)",
        "([一-龯々〆ヵヶ]{1,4}[ 　]+[一-龯々〆ヵヶ]{1,5})\s*(?:様|さん|殿|氏)"
    )
    foreach ($pattern in $namePatterns) {
        foreach ($match in [regex]::Matches($Text, $pattern)) {
            $value = if ($match.Groups.Count -gt 1) { $match.Groups[1].Value.Trim() } else { $match.Value.Trim() }
            if ($value.Length -ge 2 -and $value.Length -le 20) {
                [void]$items.Add([pscustomobject]@{
                    Type = "Name"
                    Value = $value
                    Reason = "自由記述欄の氏名らしき文字列"
                })
            }
        }
    }

    return $items
}

function Add-Finding {
    param(
        [string]$Sheet,
        [string]$Cell,
        [string]$Type,
        [string]$Kind,
        [string]$Original,
        [string]$Replacement,
        [string]$Reason
    )
    if ([string]::IsNullOrWhiteSpace($Original)) { return }
    $dedupeKey = "$Sheet|$Cell|$Type|$Kind|$Original"
    foreach ($existing in $script:Findings) {
        if ($existing.DedupeKey -eq $dedupeKey) { return }
    }
    [void]$script:Findings.Add([pscustomobject]@{
        Use = $true
        Sheet = $Sheet
        Cell = $Cell
        Type = $Type
        Kind = $Kind
        Original = $Original
        Replacement = $Replacement
        Reason = $Reason
        DedupeKey = $dedupeKey
    })
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function New-TempWorkbookCopy {
    Clear-TempWorkbook
    $script:TempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ExcelPrivacyCleaner_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $script:TempDir | Out-Null
    $target = Join-Path $script:TempDir ([System.IO.Path]::GetFileName($script:SelectedFile))
    Copy-Item -LiteralPath $script:SelectedFile -Destination $target -Force
    $script:TempWorkbook = $target
}

function Scan-Workbook {
    param([System.Windows.Forms.DataGridView]$Grid, [System.Windows.Forms.ToolStripStatusLabel]$Status)

    Reset-State
    New-TempWorkbookCopy
    $Status.Text = "検査中: Excel をローカルで読み取っています..."
    [System.Windows.Forms.Application]::DoEvents()

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($script:TempWorkbook, $null, $true)

        foreach ($sheet in $workbook.Worksheets) {
            $used = $sheet.UsedRange
            $rowCount = [int]$used.Rows.Count
            $colCount = [int]$used.Columns.Count
            if ($rowCount -lt 1 -or $colCount -lt 1) { continue }

            $columnKinds = @{}
            $headerScanRows = [Math]::Min($rowCount, 10)
            for ($row = 1; $row -le $headerScanRows; $row++) {
                for ($col = 1; $col -le $colCount; $col++) {
                    $header = Get-SafeText (($sheet.Cells.Item($row, $col)).Text)
                    if (Test-NameHeader $header) {
                        $columnKinds[$col] = @{ Type = "Name"; HeaderRow = $row; Header = $header }
                    } elseif (Test-AddressHeader $header) {
                        $columnKinds[$col] = @{ Type = "Address"; HeaderRow = $row; Header = $header }
                    }
                }
            }

            for ($row = 1; $row -le $rowCount; $row++) {
                for ($col = 1; $col -le $colCount; $col++) {
                    $cell = $sheet.Cells.Item($row, $col)
                    $text = Get-SafeText $cell.Text
                    if ([string]::IsNullOrWhiteSpace($text)) { continue }

                    $address = $cell.Address($false, $false)
                    if ($columnKinds.ContainsKey($col) -and $row -gt [int]$columnKinds[$col].HeaderRow) {
                        $type = $columnKinds[$col].Type
                        $replacement = Get-Alias -Type $type -Original $text
                        Add-Finding -Sheet $sheet.Name -Cell $address -Type $type -Kind "列単位" -Original $text -Replacement $replacement -Reason "見出し「$($columnKinds[$col].Header)」"
                    }

                    $candidates = Get-DetectionCandidates -Text $text
                    foreach ($candidate in $candidates) {
                        $replacement = Get-Alias -Type $candidate.Type -Original $candidate.Value
                        Add-Finding -Sheet $sheet.Name -Cell $address -Type $candidate.Type -Kind "自由記述" -Original $candidate.Value -Replacement $replacement -Reason $candidate.Reason
                    }
                }
            }
        }
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    $Grid.Rows.Clear()
    foreach ($finding in $script:Findings) {
        [void]$Grid.Rows.Add($finding.Use, $finding.Sheet, $finding.Cell, $finding.Type, $finding.Kind, $finding.Original, $finding.Replacement, $finding.Reason)
    }
    $Status.Text = "検査完了: $($script:Findings.Count) 件を検出しました。変換対象を確認してください。"
}

function Get-OutputPath {
    $folder = [System.IO.Path]::GetDirectoryName($script:SelectedFile)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:SelectedFile)
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $candidate = Join-Path $folder ("{0}_匿名化_{1}.xlsx" -f $baseName, $stamp)
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    return Join-Path $folder ("{0}_匿名化_{1}_{2}.xlsx" -f $baseName, $stamp, [guid]::NewGuid().ToString("N").Substring(0, 6))
}

function Convert-Workbook {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [System.Windows.Forms.ListBox]$History,
        [System.Windows.Forms.ToolStripStatusLabel]$Status
    )

    if (-not $script:TempWorkbook -or -not (Test-Path -LiteralPath $script:TempWorkbook)) {
        throw "先に検査を実行してください。"
    }

    $rows = @()
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        if ([bool]$row.Cells["Use"].Value) {
            $rows += [pscustomobject]@{
                Sheet = [string]$row.Cells["Sheet"].Value
                Cell = [string]$row.Cells["Cell"].Value
                Type = [string]$row.Cells["Type"].Value
                Kind = [string]$row.Cells["Kind"].Value
                Original = [string]$row.Cells["Original"].Value
                Replacement = [string]$row.Cells["Replacement"].Value
            }
        }
    }

    if ($rows.Count -eq 0) {
        throw "変換対象が選択されていません。"
    }

    $Status.Text = "変換中: 一時コピーへ置換を適用しています..."
    [System.Windows.Forms.Application]::DoEvents()

    $outputPath = Get-OutputPath
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($script:TempWorkbook, $null, $false)

        $sheetCache = @{}
        foreach ($sheet in $workbook.Worksheets) {
            $sheetCache[$sheet.Name] = $sheet
        }

        foreach ($item in $rows) {
            if (-not $sheetCache.ContainsKey($item.Sheet)) { continue }
            $cell = $sheetCache[$item.Sheet].Range($item.Cell)
            $current = Get-SafeText $cell.Text
            if ($item.Kind -eq "列単位") {
                $cell.Value2 = $item.Replacement
            } else {
                $cell.Value2 = $current.Replace($item.Original, $item.Replacement)
            }
        }

        $workbook.SaveAs($outputPath, 51)
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Clear-TempWorkbook
    }

    $message = "{0}  {1} 件変換  {2} -> {3}  一時ファイル削除済み" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $rows.Count, [System.IO.Path]::GetFileName($script:SelectedFile), [System.IO.Path]::GetFileName($outputPath)
    [void]$History.Items.Insert(0, $message)
    $Status.Text = "保存完了: $outputPath"
    [System.Windows.Forms.MessageBox]::Show(
        "匿名化済み Excel を保存しました。`n`n$outputPath`n`n原本は上書きしていません。一時コピーは削除済みです。",
        "保存完了",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function New-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "hoso Privacy Cleaner"
    $form.Width = 1180
    $form.Height = 760
    $form.MinimumSize = New-Object System.Drawing.Size(980, 640)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
    $form.AllowDrop = $true

    $main = New-Object System.Windows.Forms.TableLayoutPanel
    $main.Dock = "Fill"
    $main.RowCount = 5
    $main.ColumnCount = 1
    $main.Padding = New-Object System.Windows.Forms.Padding(12)
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 78))) | Out-Null
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42))) | Out-Null
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 70))) | Out-Null
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30))) | Out-Null
    $main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null

    $header = New-Object System.Windows.Forms.Label
    $header.Dock = "Fill"
    $header.Text = "Excel ファイルを選択またはドラッグ&ドロップしてください。処理はこの PC 内だけで行い、原本は上書きしません。"
    $header.TextAlign = "MiddleLeft"
    $header.Font = New-Object System.Drawing.Font("Meiryo UI", 11, [System.Drawing.FontStyle]::Bold)

    $controls = New-Object System.Windows.Forms.TableLayoutPanel
    $controls.Dock = "Fill"
    $controls.ColumnCount = 5
    $controls.RowCount = 1
    $controls.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    foreach ($width in @(112, 112, 180, 112)) {
        $controls.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, $width))) | Out-Null
    }

    $pathBox = New-Object System.Windows.Forms.TextBox
    $pathBox.Dock = "Fill"
    $pathBox.ReadOnly = $true

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = "Excelを選択"
    $selectButton.Dock = "Fill"

    $scanButton = New-Object System.Windows.Forms.Button
    $scanButton.Text = "検査開始"
    $scanButton.Dock = "Fill"

    $convertButton = New-Object System.Windows.Forms.Button
    $convertButton.Text = "確認済みを変換保存"
    $convertButton.Dock = "Fill"

    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "履歴消去"
    $clearButton.Dock = "Fill"

    $controls.Controls.Add($pathBox, 0, 0)
    $controls.Controls.Add($selectButton, 1, 0)
    $controls.Controls.Add($scanButton, 2, 0)
    $controls.Controls.Add($convertButton, 3, 0)
    $controls.Controls.Add($clearButton, 4, 0)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $true

    $colUse = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colUse.Name = "Use"
    $colUse.HeaderText = "変換"
    $colUse.FillWeight = 40
    [void]$grid.Columns.Add($colUse)

    foreach ($column in @(
        @{ Name = "Sheet"; Header = "シート"; Weight = 75 },
        @{ Name = "Cell"; Header = "セル"; Weight = 45 },
        @{ Name = "Type"; Header = "種類"; Weight = 55 },
        @{ Name = "Kind"; Header = "検査"; Weight = 65 },
        @{ Name = "Original"; Header = "検出値"; Weight = 170 },
        @{ Name = "Replacement"; Header = "変換後"; Weight = 100 },
        @{ Name = "Reason"; Header = "理由"; Weight = 150 }
    )) {
        $textCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $textCol.Name = $column.Name
        $textCol.HeaderText = $column.Header
        $textCol.FillWeight = $column.Weight
        if ($column.Name -notin @("Replacement")) { $textCol.ReadOnly = $true }
        [void]$grid.Columns.Add($textCol)
    }

    $historyLabel = New-Object System.Windows.Forms.Label
    $historyLabel.Text = "変換履歴"
    $historyLabel.Dock = "Top"
    $historyLabel.Height = 24
    $historyLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 9, [System.Drawing.FontStyle]::Bold)

    $history = New-Object System.Windows.Forms.ListBox
    $history.Dock = "Fill"

    $historyPanel = New-Object System.Windows.Forms.Panel
    $historyPanel.Dock = "Fill"
    $historyPanel.Controls.Add($history)
    $historyPanel.Controls.Add($historyLabel)

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $status = New-Object System.Windows.Forms.ToolStripStatusLabel
    $status.Text = "待機中: 外部クラウドへ送信しません。"
    [void]$statusStrip.Items.Add($status)

    $main.Controls.Add($header, 0, 0)
    $main.Controls.Add($controls, 0, 1)
    $main.Controls.Add($grid, 0, 2)
    $main.Controls.Add($historyPanel, 0, 3)
    $main.Controls.Add($statusStrip, 0, 4)
    $form.Controls.Add($main)

    $setFile = {
        param([string]$file)
        if (-not $file) { return }
        if ($file -notmatch "\.(xlsx|xlsm|xls)$") {
            [System.Windows.Forms.MessageBox]::Show("Excel ファイル（.xlsx / .xlsm / .xls）を選択してください。", "形式エラー") | Out-Null
            return
        }
        Clear-TempWorkbook
        Reset-State
        $grid.Rows.Clear()
        $script:SelectedFile = $file
        $pathBox.Text = $file
        $status.Text = "選択済み: 検査開始を押してください。"
    }

    $selectButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
        $dialog.Title = "検査する Excel を選択"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            & $setFile $dialog.FileName
        }
    })

    $scanButton.Add_Click({
        try {
            if (-not $script:SelectedFile) { throw "Excel ファイルを選択してください。" }
            Scan-Workbook -Grid $grid -Status $status
        } catch {
            $status.Text = "検査エラー"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "検査エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    $convertButton.Add_Click({
        try {
            Convert-Workbook -Grid $grid -History $history -Status $status
        } catch {
            $status.Text = "変換エラー"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "変換エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    $clearButton.Add_Click({
        $history.Items.Clear()
        $status.Text = "履歴を消去しました。"
    })

    $form.Add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    })

    $form.Add_DragDrop({
        $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0) {
            & $setFile $files[0]
        }
    })

    $form.Add_FormClosing({
        Clear-TempWorkbook
    })

    return $form
}

[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-MainForm
[System.Windows.Forms.Application]::Run($form)
