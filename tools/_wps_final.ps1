param([string]$Src, [string]$OutLog, [string]$Pdf)
$ErrorActionPreference = 'Stop'
$log = New-Object System.Collections.Generic.List[string]
try {
  $copy = Join-Path $env:TEMP ("final_" + [guid]::NewGuid().ToString("N").Substring(0,8) + ".docx")
  Copy-Item -LiteralPath $Src -Destination $copy
  $app = New-Object -ComObject KWps.Application
  $app.Visible = $false
  $doc = $app.Documents.Open($copy)
  $doc.Fields.Update() | Out-Null
  $log.Add("FIELDS.UPDATE ok, fields=$($doc.Fields.Count)")

  # 逐段找 Heading 1，读文本（此时"隐藏文字"按默认关闭）
  $h1 = @(); $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    if ($i -gt 2000) { break }
    if ($p.Style.NameLocal -match '^标题 ?1$|^Heading ?1$') { $h1 += $p.Range.Text.Trim() }
  }
  $log.Add("H1(默认隐藏关闭) 共 $($h1.Count) 条: " + (($h1 | Select-Object -First 8) -join ' | '))

  # 打开"显示隐藏文字"后再读
  try {
    $doc.ActiveWindow.View.ShowHiddenText = $true
    $app.ScreenRefresh()
    $h1b = @(); $i = 0
    foreach ($p in $doc.Paragraphs) {
      $i++; if ($i -gt 2000) { break }
      if ($p.Style.NameLocal -match '^标题 ?1$|^Heading ?1$') { $h1b += $p.Range.Text.Trim() }
    }
    $log.Add("H1(显示隐藏文字) 共 $($h1b.Count) 条: " + (($h1b | Select-Object -First 8) -join ' | '))
    $same = ($h1 -join '#') -eq ($h1b -join '#')
    $log.Add("两者一致: $same")
    $doc.ActiveWindow.View.ShowHiddenText = $false
  } catch {
    $log.Add("SKIP ShowHiddenText: " + $_.Exception.Message)
  }

  # 导出 PDF
  try {
    if ($Pdf -and (Test-Path -LiteralPath (Split-Path -Parent $Pdf))) {
      $doc.ExportAsFixedFormat($Pdf, 17) | Out-Null
      $log.Add("PDF 导出: " + (Test-Path -LiteralPath $Pdf) + " size=" + (Get-Item -LiteralPath $Pdf).Length)
    }
  } catch { $log.Add("PDF 失败: " + $_.Exception.Message) }

  $doc.Close(0) | Out-Null
  $app.Quit() | Out-Null
  $log.Add("DONE")
} catch {
  $log.Add("ERR: " + $_.Exception.Message)
} finally {
  [IO.File]::WriteAllText($OutLog, ($log -join "`r`n"), (New-Object Text.UTF8Encoding $false))
}
