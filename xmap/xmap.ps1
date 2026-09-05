# xmap.ps1 - Generic Excel schema-map tool (peek / extract / verify / aggregate + spec reuse)
# Usage:
#   xmap sheets <file.xlsx>
#   xmap headers <file.xlsx> -Sheet <name> [-Rows 5]
#   xmap extract <file.xlsx> -Sheet <name> [-Spec <spec.json>] -Out <records.jsonl>   (spec可选，缺省按指纹从 specs/ 自动匹配)
#   xmap verify  <file.xlsx> -Sheet <name> -Spec <spec.json> -Records <records.jsonl> [-Report report.txt]
#   xmap hash    <file.xlsx> -Sheet <name>            # header fingerprint（spec 复用）
#   xmap save-spec <file.xlsx> -Sheet <name> -Spec <spec.json>   # 按指纹存入 specs/ 目录
#   xmap aggregate -Records a.jsonl,b.jsonl -GroupBy name,category -Sum bal_1y,qty_1y -Out agg.jsonl
param(
  [Parameter(Position=0)][string]$Command = '',
  [Parameter(Position=1)][string]$File = '',
  [string]$Sheet = '',
  [string]$Spec = '',
  [string]$Out = '',
  [string]$Records = '',
  [string]$Report = '',
  [string]$SpecDir = '',
  [string]$GroupBy = '',
  [string]$Sum = '',
  [int]$Rows = 5
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function New-Excel {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  return $excel
}
function Close-Excel($excel, $wb) {
  if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
  if ($null -ne $excel) { try { $excel.Quit() } catch {} }
}
function Resolve-SpecDir {
  if ($SpecDir) { return $SpecDir }
  return (Join-Path $PSScriptRoot 'specs')
}
function LetterToCol($s) {
  $n = 0; foreach ($ch in $s.ToCharArray()) { $n = $n*26 + ([int]$ch - 64) }; return $n
}
function Resolve-Col($ws, $headerRows, $ref) {
  if ($null -eq $ref -or [string]::IsNullOrWhiteSpace([string]$ref)) { return $null }
  $rs = ([string]$ref).Trim()
  if ($rs -match '^[0-9]+$') { return [int]$rs }
  if ($rs -match '^[A-Za-z]{1,3}$') { return LetterToCol $rs.ToUpper() }
  $ncols = $ws.UsedRange.Columns.Count
  for ($c=1; $c -le $ncols; $c++) {
    for ($r=1; $r -le $headerRows; $r++) {
      $t = [string]$ws.Cells.Item($r,$c).Text
      if ($t.Trim() -eq $rs) { return $c }
    }
  }
  return $null
}
function Get-Value($cell) {
  try { $v = $cell.Value2 } catch { return $null }
  if ($null -eq $v) { return $null }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return $v }
  return [string]$v
}
function Cell-Text($cell, $maxLen) {
  try { $t = [string]$cell.Text } catch { return '' }
  $t = $t.Trim()
  if ($maxLen -gt 0 -and $t.Length -gt $maxLen) { return $t.Substring(0,$maxLen) + '...' }
  return $t
}
function Compute-Hash($excel, $sheetName) {
  $ws = $excel.Worksheets.Item($sheetName)
  $u = $ws.UsedRange
  $maxC = $u.Columns.Count
  $parts = @()
  $maxR = [Math]::Min(4, $u.Rows.Count)
  for ($r=1; $r -le $maxR; $r++) {
    $rowParts = @()
    for ($c=1; $c -le $maxC; $c++) { $rowParts += (Cell-Text $ws.Cells.Item($r,$c) 100).ToLower() }
    $parts += ($rowParts -join '~')
  }
  $joined = $parts -join '|'
  $md5 = [System.Security.Cryptography.MD5]::Create()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($joined)
  return [BitConverter]::ToString($md5.ComputeHash($bytes)).Replace('-','').ToLower()
}

function Cmd-Sheets {
  if (-not (Test-Path $File)) { throw "file not found: $File" }
  $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
  try {
    foreach ($ws in $wb.Worksheets) {
      $u = $ws.UsedRange
      Write-Output ("{0}`trows={1}`tcols={2}" -f $ws.Name, $u.Rows.Count, $u.Columns.Count)
    }
  } finally { Close-Excel $excel $wb }
}

function Cmd-Headers {
  if (-not (Test-Path $File)) { throw "file not found: $File" }
  $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
  try {
    $ws = $wb.Worksheets.Item($Sheet)
    $u = $ws.UsedRange
    $maxC = [Math]::Min($u.Columns.Count, 30)
    for ($r=1; $r -le $Rows; $r++) {
      $line = @()
      for ($c=1; $c -le $maxC; $c++) { $line += (Cell-Text $ws.Cells.Item($r,$c) 24) }
      $lineStr = $line -join ' | '
      Write-Output ("R" + $r + ": " + $lineStr)
    }
  } finally { Close-Excel $excel $wb }
}

function Cmd-Hash {
  if (-not (Test-Path $File)) { throw "file not found: $File" }
  $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
  try { Write-Output (Compute-Hash $excel $Sheet) } finally { Close-Excel $excel $wb }
}

function Cmd-SaveSpec {
  if (-not (Test-Path $File)) { throw "file not found: $File" }
  if ([string]::IsNullOrWhiteSpace($Spec)) { throw "-Spec required" }
  if (-not (Test-Path $Spec)) { throw "spec not found: $Spec" }
  $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
  $h = ''
  try { $h = Compute-Hash $excel $Sheet } finally { Close-Excel $excel $wb }
  $dir = Resolve-SpecDir
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $dest = Join-Path $dir ($h + '.json')
  Copy-Item $Spec $dest -Force
  Write-Output ("saved spec -> " + $dest)
}

function Cmd-Extract {
  if (-not (Test-Path $File)) { throw "file not found: $File" }
  if ([string]::IsNullOrWhiteSpace($Spec) -or -not (Test-Path $Spec)) {
    $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
    $h = ''
    try { $h = Compute-Hash $excel $Sheet } finally { Close-Excel $excel $wb }
    $cand = Join-Path (Resolve-SpecDir) ($h + '.json')
    if (Test-Path $cand) { $Spec = $cand; Write-Output ("auto-resolved spec by hash " + $h) }
    else { throw ("no spec for sheet; write one and run: xmap save-spec -Spec <spec>  (hash=" + $h + ")") }
  }
  $spec = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json
  $headerRows = [int]$spec.headerRows
  $dataStart = [int]$spec.dataStart
  $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
  try {
    $ws = $wb.Worksheets.Item($Sheet)
    $colMap = @{}
    foreach ($prop in $spec.cols.PSObject.Properties) {
      $colMap[$prop.Name] = Resolve-Col $ws $headerRows $prop.Value
    }
    $lastRow = $ws.UsedRange.Rows.Count
    $sb = New-Object System.Text.StringBuilder
    for ($r = $dataStart; $r -le $lastRow; $r++) {
      $rec = [ordered]@{}
      $allNull = $true
      foreach ($k in $colMap.Keys) {
        $ci = $colMap[$k]
        if ($null -eq $ci) { $rec[$k] = $null; continue }
        $v = Get-Value $ws.Cells.Item($r,$ci)
        $rec[$k] = $v
        if ($null -ne $v -and ([string]$v).Trim() -ne '') { $allNull = $false }
      }
      if ($allNull) { continue }
      if ($spec.skipIfNameEmpty -and $colMap.ContainsKey('name')) {
        $nv = $rec['name']
        if ($null -eq $nv -or ([string]$nv).Trim() -eq '') { continue }
      }
      [void]$sb.AppendLine(($rec | ConvertTo-Json -Compress -Depth 5))
    }
    Set-Content -Path $Out -Value $sb.ToString() -Encoding UTF8
    Write-Output ("records -> " + $Out)
  } finally { Close-Excel $excel $wb }
}

function Cmd-Verify {
  if ([string]::IsNullOrWhiteSpace($Spec)) { throw "-Spec required" }
  if (-not (Test-Path $Spec)) { throw "spec not found: $Spec" }
  if (-not (Test-Path $Records)) { throw "records not found: $Records" }
  $spec = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json
  $lines = Get-Content -Encoding UTF8 $Records
  $records = @(); foreach ($ln in $lines) { if ($ln.Trim()) { $records += ($ln | ConvertFrom-Json) } }
  $numCols = @(); foreach ($p in $spec.numeric.PSObject.Properties) { $numCols += $p.Name }
  $rep = New-Object System.Text.StringBuilder
  [void]$rep.AppendLine("records=" + $records.Count)
  foreach ($nc in $numCols) {
    $sum = 0.0
    foreach ($rec in $records) { if ($null -ne $rec.$nc) { $sum += [double]$rec.$nc } }
    [void]$rep.AppendLine("sum[$nc]=" + [math]::Round($sum,4))
  }
  if ($null -ne $spec.checkTotal) {
    $excel = New-Excel; $wb = $excel.Workbooks.Open($File)
    try {
      $ws = $wb.Worksheets.Item($Sheet)
      $label = [string]$spec.checkTotal.label
      $srcCol = Resolve-Col $ws $spec.headerRows $spec.checkTotal.col
      $nameCol = Resolve-Col $ws $spec.headerRows $spec.checkTotal.nameCol
      $val = $null
      $lastRow = $ws.UsedRange.Rows.Count
      for ($r = $spec.dataStart; $r -le $lastRow; $r++) {
        $name = [string]$ws.Cells.Item($r,$nameCol).Text
        if ($name.Trim() -eq $label) { $val = Get-Value $ws.Cells.Item($r,$srcCol); break }
      }
      [void]$rep.AppendLine("checkTotal[$label]=" + $val)
    } finally { Close-Excel $excel $wb }
  }
  if ($Report) { Set-Content -Path $Report -Value $rep.ToString() -Encoding UTF8; Write-Output ("report -> " + $Report) }
  Write-Output $rep.ToString()
}

function Cmd-Aggregate {
  if (-not $Records) { throw "-Records required (comma-separated jsonl files)" }
  if (-not $GroupBy) { throw "-GroupBy required (comma-separated field names)" }
  $groupKeys = @(); foreach ($k in ($GroupBy -split ',')) { $groupKeys += $k.Trim() }
  $sumKeys = @(); foreach ($k in ($Sum -split ',')) { if ($k.Trim()) { $sumKeys += $k.Trim() } }
  $agg = @{}
  foreach ($rf in ($Records -split ',')) {
    $rf = $rf.Trim()
    if (-not (Test-Path $rf)) { throw "records not found: $rf" }
    foreach ($ln in (Get-Content -Encoding UTF8 $rf)) {
      if (-not $ln.Trim()) { continue }
      $rec = $ln | ConvertFrom-Json
      $parts = @(); foreach ($gk in $groupKeys) { $parts += [string]$rec.$gk }
      $key = $parts -join '|'
      if (-not $agg.ContainsKey($key)) {
        $g = @{}; foreach ($gk in $groupKeys) { $g[$gk] = $rec.$gk }
        $sum = @{}; foreach ($sk in $sumKeys) { $sum[$sk] = 0.0 }
        $agg[$key] = @{ g=$g; s=$sum }
      }
      $e = $agg[$key]
      foreach ($sk in $sumKeys) { $v = $rec.$sk; if ($null -ne $v) { $e.s[$sk] += [double]$v } }
    }
  }
  $sb = New-Object System.Text.StringBuilder
  foreach ($k in $agg.Keys) {
    $e = $agg[$k]
    $o = [ordered]@{}
    foreach ($gk in $groupKeys) { $o[$gk] = $e.g[$gk] }
    foreach ($sk in $sumKeys) { $o[$sk] = [math]::Round($e.s[$sk],4) }
    [void]$sb.AppendLine(($o | ConvertTo-Json -Compress -Depth 5))
  }
  if ($Out) { Set-Content -Path $Out -Value $sb.ToString() -Encoding UTF8; Write-Output ("agg -> " + $Out) }
  else { Write-Output $sb.ToString() }
}

switch ($Command) {
  'sheets'   { Cmd-Sheets }
  'headers'  { Cmd-Headers }
  'extract'  { Cmd-Extract }
  'verify'   { Cmd-Verify }
  'hash'     { Cmd-Hash }
  'save-spec'{ Cmd-SaveSpec }
  'aggregate'{ Cmd-Aggregate }
  default    { Write-Output "commands: sheets | headers | extract | verify | hash | save-spec | aggregate"; exit 1 }
}