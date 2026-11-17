param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$OutLog
)
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP ('wps_fieldtest_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$copy = Join-Path $tmp 'probe.docx'
Copy-Item -LiteralPath $Src -Destination $copy -Force

$lines = New-Object System.Collections.Generic.List[string]
try {
  $app = New-Object -ComObject KWps.Application
  $app.Visible = $false
  try { $app.DisplayAlerts = 0 } catch {}
  $doc = $app.Documents.Open($copy, $false, $false)
  $lines.Add('OPENED ok')
  $before = $doc.Fields.Count
  $lines.Add("Fields.Count(before)=$before")
  $doc.Fields.Update() | Out-Null
  $lines.Add('FIELDS.UPDATE done')
  # 遍历段落，抓题注段
  $n = $doc.Paragraphs.Count
  $lines.Add("Paragraphs=$n")
  for ($i = 1; $i -le $n; $i++) {
    $t = $doc.Paragraphs.Item($i).Range.Text
    if ($null -eq $t) { continue }
    $t = $t -replace "`r", '' -replace "`a", ''
    if ($t -match '^\s*(表|图|式)\s') {
      $lines.Add("CAP[$i]=" + $t)
    }
  }
  # 也统计含"错误"字样的段落
  for ($i = 1; $i -le $n; $i++) {
    $t = $doc.Paragraphs.Item($i).Range.Text
    if ($null -eq $t) { continue }
    if ($t -match '错误|未定义') {
      $lines.Add("ERR[$i]=" + ($t -replace "`r", ''))
    }
  }
  $doc.Close(0)
  $app.Quit()
} catch {
  $lines.Add('EXCEPTION: ' + $_.Exception.Message)
  try { $doc.Close(0) } catch {}
  try { $app.Quit() } catch {}
} finally {
  foreach ($f in Get-ChildItem -LiteralPath $tmp) { try { $f.Delete() } catch {} }
  try { Remove-Item -LiteralPath $tmp -Force } catch {}
}
[IO.File]::WriteAllLines($OutLog, $lines, (New-Object Text.UTF8Encoding($false)))
Write-Output ("lines=" + $lines.Count + " -> " + $OutLog)
