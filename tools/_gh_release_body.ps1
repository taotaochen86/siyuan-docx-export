# 更新 v0.3.4 Release 说明，补入深度对账与稳定性实测结论
$ErrorActionPreference = 'Stop'
$repo = 'taotaochen86/siyuan-docx-export'
$tag = 'v0.3.4'
$proxy = 'http://127.0.0.1:7890'

$cred = "protocol=https`nhost=github.com`n`n" | git credential fill 2>$null
$tok = (($cred | Where-Object { $_ -like 'password=*' }) -replace 'password=', '').Trim()
if (-not $tok) { throw 'no token' }

$relId = 448708781
$body = @"
## 修复
- 题注/公式编号 F9 刷新不再报「错误！未定义样式。」：章号改为静态文本（标题编号本身是静态文本，STYLEREF 的 \s/\n 取不到段落编号），序号仍由 ``SEQ 表|图|式 \* ARABIC \s 1`` 按 Heading 1 每章重置。
- 交叉引用不再报「错误！未定义书签。」：题注书签改为 ``bkmk_cap_<tbl|fig|eq>_<编号>`` 并全局去重，同号多文档不再撞名；引用命中不了题注时降级为普通文本。
- 公式仍为 Word 原生 OMML 对象（含分段函数 matrix/cases、括号定界组、上下标）。

## 实测证据（WPS 12.0 COM，Fields.Update）
### 大稿（26 篇 / 46 表 / 33 图 / 151 标题 / 177 代码块）
- 深度对账 15 项全通过：标题文本 100% 命中、表格 46/46、总行数 553/553、题注 11/11、书签 11 个唯一、域 begin/end 17/17、图片 33 引用/29 媒体、分页 25、无字面转义残留。
- 域更新：17 个域全部正常，REF 输出 ``图 1-1 / 表 2-1 / 表 2-2 / 表 3-1 / 图 3-1 / (4-1)``，SEQ 每章从 1 重置，0 条错误。
- 稳定性：连续更新 2 次 + 保存后重开第 3 次，结果完全一致，0 错误。

### 测试书（8 篇 / 5 表 / 11 条 Zotero）
- 深度对账 15 项全通过：标题全命中、表格 5/5、行 17/17、题注 5/5、书签 16（题注 5 + 文献 11）、域 24/24。
- 域更新：REF ``表 1-1 / 表 4-1 / 表 4-2 / 表 4-3``，ZOTERO_ITEM 13 + ZOTERO_BIBL 1 完整刷新，0 错误。
- 稳定性：三次更新（含保存重开）结果一致，0 错误。

## 已知限制
- 标题编号为静态文本（与思源一致），不随 Word 自动编号重排；因此题注章号同样静态，不随标题改动联动。
- 引用指向「无题注的表格块」时降级为普通文本（不可跳转），以保证 F9 不出现「错误！未定义书签。」。
- 公式 OMML 子集未覆盖大型积分/求和上下限、多列真矩阵、多行对齐，超出则回退线性文本。
"@

$payload = @{ body = $body } | ConvertTo-Json -Depth 3
$resp = curl.exe -s -x $proxy -X PATCH "https://api.github.com/repos/$repo/releases/$relId" `
  -H "Authorization: token $tok" -H "Accept: application/vnd.github+json" `
  -H "Content-Type: application/json" --data-binary $payload
$r = $resp | ConvertFrom-Json
Write-Output "patched release id=$($r.id) url=$($r.html_url)"
