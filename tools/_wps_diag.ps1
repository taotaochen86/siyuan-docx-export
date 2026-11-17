param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$OutLog
)
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP ('wps_diag_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$copy = Join-Path $tmp 'probe.docx'
Copy-Item -LiteralPath $Src -Destination $copy -Force

$L = New-Object System.Collections.Generic.List[string]
$doc = $null; $app = $null
try {
  $app = New-Object -ComObject KWps.Application
  $app.Visible = $false
  try { $app.DisplayAlerts = 0 } catch {}
  $doc = $app.Documents.Open($copy, $false, $false)

  $L.Add('=== 书签 ===')
  $L.Add('Bookmarks=' + $doc.Bookmarks.Count)
  $nbm = $doc.Bookmarks.Count
  for ($i = 1; $i -le $nbm; $i++) {
    $L.Add('BM[' + $i + ']=' + $doc.Bookmarks.Item($i).Name)
  }

  $L.Add('=== 域 ===')
  $L.Add('Fields=' + $doc.Fields.Count)
  $nf = $doc.Fields.Count
  $refShown = 0
  $stShown = 0
  for ($i = 1; $i -le $nf; $i++) {
    $c = $doc.Fields.Item($i).Code.Text
    $r = $doc.Fields.Item($i).Result.Text
    if ($c -match 'REF ' -and $refShown -lt 12) {
      $L.Add("REF[$i] code=[$c] result=[$r]"); $refShown++
    }
    if ($c -match 'STYLEREF' -and $stShown -lt 6) {
      $L.Add("STY[$i] code=[$c] result=[$r]"); $stShown++
    }
    if ($c -match 'SEQ ') {
      $L.Add("SEQ[$i] code=[$c] result=[$r]")
    }
  }

  # 找第一个 STYLEREF 域做写法实验
  $tgt = $null
  for ($i = 1; $i -le $nf; $i++) {
    if ($doc.Fields.Item($i).Code.Text -match 'STYLEREF') { $tgt = $doc.Fields.Item($i); break }
  }
  $L.Add('=== STYLEREF 写法实验 ===')
  if ($null -eq $tgt) { $L.Add('no STYLEREF field') }
  else {
    $variants = @(
      ' STYLEREF "Heading 1" \s ',
      ' STYLEREF "heading 1" \s ',
      ' STYLEREF "标题 1" \s ',
      ' STYLEREF "Heading 1" ',
      ' STYLEREF "heading 1" ',
      ' STYLEREF "标题 1" ',
      ' STYLEREF "heading 1" \n ',
      ' STYLEREF "标题 1" \n '
    )
    foreach ($v in $variants) {
      try {
        $tgt.Code.Text = $v
        $tgt.Update() | Out-Null
        $res = $tgt.Result.Text
        $L.Add("[$v] => [$res]")
      } catch {
        $L.Add("[$v] => EXC " + $_.Exception.Message)
      }
    }
  }

  $doc.Close(0)
  $app.Quit()
  $doc = $null; $app = $null
} catch {
  $L.Add('EXCEPTION: ' + $_.Exception.Message)
  try { if ($doc) { $doc.Close(0) } } catch {}
  try { if ($app) { $app.Quit() } } catch {}
} finally {
  Start-Sleep -Milliseconds 300
  try { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction Stop } catch { $L.Add('tmp cleanup failed: ' + $_.Exception.Message) }
}
[IO.File]::WriteAllLines($OutLog, $L, (New-Object Text.UTF8Encoding($false)))
Write-Output ('lines=' + $L.Count + ' -> ' + $OutLog)
