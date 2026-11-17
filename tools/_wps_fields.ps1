param([Parameter(Mandatory=$true)][string]$Src, [Parameter(Mandatory=$true)][string]$OutLog)
$ErrorActionPreference = 'Stop'
$lines = New-Object System.Collections.Generic.List[string]
$app = $null; $doc = $null
try {
  $app = New-Object -ComObject KWps.Application
  $app.Visible = $false; $app.DisplayAlerts = 0
  $doc = $app.Documents.Open($Src, $false, $true)
  $lines.Add("OPENED ok")
  $doc.Fields.Update() | Out-Null
  $lines.Add("FIELDS.UPDATE done count=$($doc.Fields.Count)")
  for ($i = 1; $i -le $doc.Fields.Count; $i++) {
    $f = $doc.Fields.Item($i)
    $code = ($f.Code.Text -replace "\s+", " ").Trim()
    $res  = ($f.Result.Text -replace "\s+", " ").Trim()
    $lines.Add("F[$i] code=[$code] result=[$res]")
  }
} catch {
  $lines.Add("ERR: $($_.Exception.Message)")
} finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
}
[IO.File]::WriteAllLines($OutLog, $lines, (New-Object Text.UTF8Encoding $false))
Write-Output "lines=$($lines.Count) -> $OutLog"
