# 建 Release 并上传 asset（走本机代理 7890，凭据取自 git credential fill，不落盘明文）
$ErrorActionPreference = 'Stop'
$repo = 'taotaochen86/siyuan-docx-export'
$tag = 'v0.3.4'
$zip = "$env:TEMP\siyuan-rel\v0.3.4\siyuan-docx-export.zip"
$proxy = 'http://127.0.0.1:7890'

$cred = "protocol=https`nhost=github.com`n`n" | git credential fill 2>$null
$tok = (($cred | Where-Object { $_ -like 'password=*' }) -replace 'password=', '').Trim()
if (-not $tok) { throw 'no token from credential helper' }

$body = @{
  tag_name = $tag
  name = 'v0.3.4 题注章号静态化 + 交叉引用按编号书签'
  body = @"
## 修复
- 题注/公式编号 F9 刷新不再报「错误！未定义样式。」：章号改为静态文本（标题编号本身是静态文本，STYLEREF 的 \s/\n 取不到段落编号），序号仍由 ``SEQ 表|图|式 \* ARABIC \s 1`` 按 Heading 1 每章重置。
- 交叉引用不再报「错误！未定义书签。」：题注书签改为 ``bkmk_cap_<tbl|fig|eq>_<编号>`` 并全局去重，同号多文档不再撞名；引用命中不了题注时降级为普通文本。
- 公式仍为 Word 原生 OMML 对象（含分段函数 matrix/cases、括号定界组、上下标）。

## 实测（WPS 12.0 COM，Fields.Update 后）
- 大稿 26 篇 / 46 表 / 33 图：审计 13 项通过 0 失败；17 个域全部正常，REF 输出 图 1-1 / 表 2-1 / 表 2-2 / 表 3-1 / 图 3-1 / (4-1)，零错误。
- 测试书 8 篇 / 11 条 Zotero：24 个域全部正常，ZOTERO_BIBL 书目完整刷新。
"@
  draft = $false
  prerelease = $false
} | ConvertTo-Json -Depth 5

$resp = curl.exe -s -x $proxy -X POST "https://api.github.com/repos/$repo/releases" `
  -H "Authorization: token $tok" -H "Accept: application/vnd.github+json" `
  -H "Content-Type: application/json" --data-binary $body
$rel = $resp | ConvertFrom-Json
if (-not $rel.id) { throw "create release failed: $resp" }
Write-Output "release id=$($rel.id) url=$($rel.html_url)"

$up = curl.exe -s -x $proxy -X POST "https://uploads.github.com/repos/$repo/releases/$($rel.id)/assets?name=siyuan-docx-export.zip" `
  -H "Authorization: token $tok" -H "Content-Type: application/zip" `
  --data-binary "@$zip"
$asset = $up | ConvertFrom-Json
Write-Output "asset id=$($asset.id) state=$($asset.state) size=$($asset.size) dl=$($asset.browser_download_url)"
