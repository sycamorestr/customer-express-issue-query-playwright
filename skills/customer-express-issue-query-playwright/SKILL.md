---
name: customer-express-issue-query-playwright
description: 用 Playwright + CDP 连接已登录的 Edge/Chrome，在聚水潭高阶版订单页按 Excel 的月结赔付单号或快递单号批量查询发货时间、实付金额和备注，无损回填新 Excel。适用于客服快递问题件、多工作表、多单号、多订单及断点续查。
---

# 客服快递问题件查询（Playwright + CDP）

这是聚水潭**高阶版**的只读订单查询技能。读取 Excel 的快递单号，在真实订单页 `/app/order/order/list.aspx` 的登录上下文调用 `LoadDataToJSON`。不要套用分销版商品成本或物流轨迹流程。

## 环境与连接

使用内置 Windows x64 Python 和 Playwright，无需安装 Python、Node.js、浏览器扩展或 Playwright 浏览器。浏览器由使用者提供，优先 Edge，也支持 Chrome。完整分享包额外附有 `README.md`、`check.cmd`、`install.ps1` 和依赖来源说明，这些只位于最初的完整解压目录；安装后执行查询只需本技能目录。

每次新 shell 设置本技能实际路径，所有 Python 命令使用同一个解释器：

```powershell
$skill = 'D:\Tools\customer-express-issue-query-playwright\skills\customer-express-issue-query-playwright'
$python = Join-Path $skill 'runtime\python\python.exe'
$endpoint = 'http://127.0.0.1:9222'
& $python -B -c "import openpyxl; from playwright.sync_api import sync_playwright; print('Dependencies ready')"
```

已有明确的 CDP 环境时，直接采用其本机 HTTP 端点；不得根据普通浏览器已打开就断言可连接。没有 CDP 浏览器时运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\start_browser.ps1" -Port 9222 -StartUrl 'https://www.erp321.com/epaas'
```

使用者说明浏览器已通过浏览器工作台登记时，优先复用登记。从工作台连接详情，或 `%LOCALAPPDATA%\BrowserWorkbench\settings.json` 的 `registry` 找到原清单及同目录 `.browser-workbench-shops.json`，按环境名称和登记主页确认目标，沿 `browser_config`（相对清单目录解析）读取 `remote_debugging_port`，组成本机 CDP 端点。`settings.json` 的 `port` 是工作台面板端口，不是 CDP 端口。不要扫描 Profile、读取 Cookie 或为已登记环境另建浏览器；多个目标时先消歧。

启动器使用独立目录 `%LOCALAPPDATA%\CustomerExpressIssueQueryPlaywright\browser-profile`，保留使用者自己的登录态。普通浏览器窗口不能事后直接开启 CDP；首次由使用者在这个独立窗口登录。需要验证码或人工登录时保留工作目录，待完成后续查。浏览器窗口若未显示，可从任务栏切换到新实例。不要读取、导出或索要密码、Cookie、Token。

如果使用其他专用浏览器配置，可传 `-UserDataDir` 和 `-Port`；不与日常默认配置混用。启动器不关闭已有浏览器。同一配置已由其他端口使用时，复用原端点或选择另一专用配置。Windows PowerShell 5.1 与 PowerShell 7 均可使用，进程级 ExecutionPolicy 不修改系统策略。

## 前置确认

1. 告知使用本技能，确定输入 `.xlsx` 路径和业务工作目录。新任务使用独立目录；**续查必须沿用已有 `manifest.json` 和 `results.json` 的原工作目录，跳过 prepare**，不能另建任务重查。业务文件、结果和日志不放在技能或分享包中。
2. 在 CDP 浏览器中确认登录的是本次任务需要的聚水潭**高阶版账号/店铺**，并打开订单列表。已有平台账号工具可以帮助打开目标，但最终仍须验证 CDP 连接。不要猜账号，不修改订单、备注或订单状态。
3. 执行连接检查：

```powershell
& $python -B "$skill\scripts\cdp_query.py" doctor --endpoint $endpoint
```

`connected:true` 只证明连接；必须同时有 `ok:true`、唯一的目标订单 frame 和 `ready:true`。支持顶层页面及嵌套 iframe，不依赖 iframe ID。默认只接受 HTTPS 的 `erp321.com` 及其子域的订单路径；只有已核实的其他 ERP 主机才用 `--allowed-host <精确主机名>`。

已连接但没有订单 frame 时，先查看已连接页面的实际可见入口；若仍在聚水潭首页，可打开页面上的“订单”标签或菜单，然后重新运行 doctor。不要因未打开订单页而误判为需要重建浏览器或重新登录。

多个页面有订单 frame 时，先从 doctor 的 `pageIndex` 和页面路径确认目标，再为 doctor/query 指定 `--page-index N`（batch.ps1 对应 `-PageIndex N`）。同一页面包含多个订单 frame 时，关闭无关的订单标签或让使用者保留一个目标页面，不能随便选首个。索引可能随标签页变化，断点续查前重新检查账号与页面。

## 准备 Excel

仅新任务执行本节。已有断点时，先在原 CDP 浏览器恢复登录、确认账号并运行 doctor，再直接进入“查询、断点与复查”。不要对已有断点再次 prepare。

```powershell
$work = 'D:\售后任务\本次查询'
$inputFile = 'D:\售后\快递问题件.xlsx'
& $python -B "$skill\scripts\excel.py" prepare --input $inputFile --workdir $work
```

只支持 `.xlsx`。其他格式需要可靠地转为新副本并核验；不能直接改后缀。默认处理全部工作表，在前 50 行定位精确表头：优先“月结赔付单号”，其次“快递单号/物流单号/运单号/发货单号”，这些都不存在时才识别“单号”。同一优先级有多个候选时停止，不能自行取第一列；“订单号”不自动当作快递号。prepare 摘要会显示每表采用的表头、列和处理行数。逐表核对 `skippedSheets`；快递明细页缺少已知表头不能当作说明页略过。歧义时用 `--column '月结赔付单号'` 或 `--sheets '圆通,中通'` 指定，换新工作目录 prepare。

支持同格多个单号和已知快递公司标签；查询键转大写并去掉一个 `@` 前缀。允许单号文本首尾多余的空白、英文/中文逗号、顿号、分号及竖线，清理仅作用于查询键，源单元格不变。只含分隔符、未知标点或说明文字、公式、不完整单号、疑似失去精度的数字仍进入 `prepare-errors.json`，不得猜测或静默跳过。读取实际单元格，不信任错误的 dimension 元数据。

`manifest.json` 保存源路径、SHA256、所有原行映射和去重单号。查询开始后不改清单、源文件，不把工作目录复用给另一任务。

## 查询、断点与复查

先查询一批并核对四字段、匹配数量和当前页面：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\batch.ps1" -WorkDir $work -Endpoint $endpoint -MaxBatches 1
```

确认后继续全量，相同工作目录自动跳过已完成单号：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\batch.ps1" -WorkDir $work -Endpoint $endpoint
```

也可直接使用 Python：

```powershell
& $python -B "$skill\scripts\cdp_query.py" query --workdir $work --endpoint $endpoint --batch-size 25
```

查询只在选定 frame 内异步执行 `fetch`，使用页面登录态和隐藏字段，不把凭据传到 Python。筛选 `l_id`、运算符 `@=`，同时提交原单号与 `@` 前缀形式，不添加日期筛选。每批默认 25 个，允许 1–60，批间默认 150ms，单次请求默认 30 秒超时。

每次响应达到 500 条自动拆批；单号仍达到 500 条则停止，不能当作完整结果。保留多物流字段匹配和有单号边界的全文兜底。计数日志的 `fallback > 0` 时抽核该批，不把兜底当作精确物流字段匹配。

只有完整通过校验的批次才原子保存到 `results.json`。同一工作目录有进程锁，不并发执行。接口不成功、401/403、超时、登录/验证拦截、结构异常或页面切换都停止并保留已完成断点；失败不冒充未找到。恢复后使用原参数继续。日志只含计数，结果文件仅含所需四字段、单号和匹配数，不保存完整订单对象。

首次全量完成后，以小批次复查未找到：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$skill\scripts\batch.ps1" -WorkDir $work -Endpoint $endpoint -RecheckMissing -BatchSize 5
```

## 回填规则与导出

顺序追加五列：查询快递单号、发货时间(confirm_date)、实付金额(paid_amount)、备注(remark)、查询结果。四字段空值转空串、连续空白合并、金额 0 保留。完整保留接口备注，不筛商品、不改写，也不覆盖原表“价值”列。

一个单号多个订单按接口顺序换行，四列对齐，不只取第一条、不做金额合计。一行多个单号回填原行，不插入或删除行。重复单号只查一次，所有原行均回填。只有接口正常返回空匹配才能标记“未找到订单”。

```powershell
& $python -B "$skill\scripts\excel.py" export --workdir $work
# 或指定尚不存在的新文件：
& $python -B "$skill\scripts\excel.py" export --workdir $work --output 'D:\售后\结果新文件.xlsx'
```

默认输出 `<源文件名>_高阶版订单查询结果.xlsx`，不覆盖源文件或已有结果。直接编辑 XLSX ZIP 的工作表 XML，不能改为 openpyxl/ExcelJS 全量重存。新列放在所有原单元格（含格式化空白）及合并区域之后，可能靠右；交付时明确列范围。

新增五列按顺序设置宽度 24、32、30、80、32，便于阅读常见单号和时间，备注列留出较大空间。原列宽、默认列宽和原行高保持不变；若原列定义跨到新增区域，只拆分该定义并设置新增五列，其他列的所有属性保留。多订单换行及超长备注仍完整存储，固定原行高可能无法一次显示全文，可点选单元格查看；不要为展示结果而整体调整原行高、裁剪备注或截断单号。

导出器逐项验证源 SHA256、原单元格/公式/格式/工作表结构、结果列之外的原列属性、所有未修改 ZIP 部件（图片、WPS cellimages、关系、附件等）的字节，以及重新读取的新结果单元格。新增列宽单独校验后，仅在结构比较时恢复原列定义；不能简单跳过列格式检查。遇到重叠列定义、多个列定义块、超过 32767 UTF-16 字符或非法 XML 控制字符时停止并保留断点，不自行重排或截断。未完成、源变更、坏结果不能宣称完成。

最后报告结果文件路径、各表处理行数与新增列范围、唯一单号/匹配/复查后未找到/多订单数及源文件未改。脚本退出仅断开 Playwright，保留使用者浏览器和登录态。不要调用关闭用户浏览器的命令。

## 维护

`scripts/excel.py` 负责表头识别、单号解析与无损导出；`scripts/query.js` 是已选中订单 frame 内的查询函数；`scripts/cdp_query.py` 负责连接、选择、校验、锁与断点；`scripts/start_browser.ps1` 只负责专用浏览器启动。页面变化时依实时证据最小修正，不硬编码真实单号、账号或 iframe ID。

修改 Excel 流程后可运行离线回归：`& $python -B -m unittest discover -s "$skill\tests" -p 'test_*.py' -v`。使用合成数据验证表头优先级、边界分隔符、无损保留和列宽；真实客户表和查询结果不放进技能包。已有查询结果可在独立验证目录离线复用，避免为导出检查重复访问聚水潭。
