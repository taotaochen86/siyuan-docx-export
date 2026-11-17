param([Parameter(Mandatory=$true)][string]$Src, [Parameter(Mandatory=$true)][string]$OutLog)
$ErrorActionPreference = 'Stop'
$lines = New-Object System.Collections.Generic.List[string]
$work = Join-Path $env:TEMP 'siyuan-stability'
if (-not (Test-Path -LiteralPath $work)) { New-Item -ItemType Directory -Force -Path $work | Out-Null }
$copy = Join-Path $work 'stability.docx'
Copy-Item -LiteralPath $Src -Destination $copy -Force

function Get-FieldReport($doc, [string]$tag, $lines) {
  $errs = 0; $refs = @()
  for ($i = 1; $i -le $doc.Fields.Count; $i++) {
    $f = $doc.Fields.Item($i)
    $code = ($f.Code.Text -replace '\s+', ' ').Trim()
    $res  = ($f.Result.Text -replace '\s+', ' ').Trim()
    if ($res -match '错误') { $errs++ }
    if ($code -like 'REF *') { $refs += $res }
  }
  $lines.Add("[$tag] fields=$($doc.Fields.Count) errors=$errs refs=$($refs -join ' | ')")
}

$app = $null; $doc = $null
try {
  $app = New-Object -ComObject KWps.Application
  $app.Visible = $false; $app.DisplayAlerts = 0
  $doc = $app.Documents.Open($copy, $false, $false)
  $doc.Fields.Update() | Out-Null
  Get-FieldReport $doc 'update#1' $lines
  $doc.Fields.Update() | Out-Null
  Get-FieldReport $doc 'update#2(幂等)' $lines
  $doc.Save()
  $lines.Add('saved ok')
  $doc.Close(0); $doc = $null
  $doc = $app.Documents.Open($copy, $false, $true)
  $doc.Fields.Update() | Out-Null
  Get-FieldReport $doc 'reopen+update#3' $lines
  # 抽查题注文本
  $caps = @()
  for ($p = 1; $p -le $doc.Paragraphs.Count; $p++) {
    $t = ($doc.Paragraphs.Item($p).Range.Text + '').Trim()
    if ($t -match '^(表|图|式)\s?[0-9]') { $caps += ($t -replace '\s+', ' ').Substring(0, [Math]::Min(30, $t.Length)) }
    if ($caps.Count -ge 6) { break }
  }
  $lines.Add("前 6 条题注: $($caps -join ' ||| ')")
} catch {
  $lines.Add("ERR: $($_.Exception.Message)")
} finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
}
[IO.File]::WriteAllLines($OutLog, $lines, (New-Object Text.UTF8Encoding $false))
Write-Output "lines=$($lines.Count) -> $OutLog"
