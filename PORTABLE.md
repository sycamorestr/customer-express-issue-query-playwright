# 客服快递问题件对账 Skill：Playwright + CDP Windows x64 便携包

面向客服部门的快递问题件对账。读取快递问题件 Excel，在聚水潭**高阶版**订单列表批量查询发货时间、实付金额、备注，回填到新 Excel 供逐项核对。支持多工作表、多单号、多订单、去重、断点续查、未找到复查，保留源表公式、格式及附件。

此版本使用 Python Playwright 的 `connect_over_cdp()`，不再需要 OpenCLI 或浏览器扩展。原 OpenCLI 版本可继续独立保留。

仓库统一命名为 `customer-service-express-issue-reconciliation-skill`；为兼容已有便携包，技能调用标识和目录仍使用 `customer-express-issue-query-playwright`，以下安装、查询和运行时路径继续使用此兼容标识。

## 已内置和仍需具备的环境

包内包含 Python 3.13.12 x64、Playwright 1.63.0（含自己的 Node 驱动）、openpyxl 3.1.5、全部 Python 依赖及原始 wheel。无需全局 pip/npm 安装，也无需下载 Playwright 浏览器。依赖清单及来源见 `python-dependencies.json`、`download-provenance.json` 和 `THIRD_PARTY.md`。

使用者电脑需要 Windows 10/11 x64、Edge 或 Chrome、PowerShell、可运行技能的 AI 工具，以及自己的聚水潭高阶版账号和订单访问权限。浏览器、AI 工具、业务 Excel 和登录态不在包内。

## 1. 完整解压与自检

完整解压到可写路径，例如 `D:\Tools\customer-express-issue-query-playwright`，不要在压缩软件内执行。双击 `check.cmd`，或执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\check.ps1
```

自检校验内置文件和运行时，使用虚构数据测试查询逻辑、断点和 Excel 保留规则；不打开浏览器、不访问真实订单，也不修改系统 PATH。SHA256 清单用于完整性检查，不是数字签名。组织策略若禁止运行，应由管理员处理。

## 2. 安装到 AI 工具（也可直接从解压目录使用）

DSH 工作区：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace 'D:\自己的工作区'
```

安装到其他工具的技能目录，例如 Codex：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -SkillsDir "$env:USERPROFILE\.codex\skills"
```

只安装一个新技能 `customer-express-issue-query-playwright`，连同其内置 runtime。已有同名目录会停止，不覆盖或合并；原版技能名称不同。也可手动复制 `skills` 内该完整文件夹，然后重新加载技能列表。

## 3. 启动专用浏览器并登录

双击 `start-browser.cmd`，默认优先打开 Edge 并提供 `http://127.0.0.1:9222`。需要直接打开聚水潭：

```powershell
.\start-browser.cmd -StartUrl 'https://www.erp321.com/epaas'
# Chrome 或其他端口：
.\start-browser.cmd -Browser Chrome -Port 9223 -UserDataDir 'D:\ERP浏览器\高阶版'
```

在新浏览器窗口登录自己的聚水潭高阶版，打开“订单”列表。第一次登录由本人完成；后续会保存在专用目录 `%LOCALAPPDATA%\CustomerExpressIssueQueryPlaywright\browser-profile`。如果窗口未显示，从任务栏切换到新实例。运行结束不会关闭浏览器。

原来日常浏览器里的登录不会自动搬过来。启动器不会使用 Edge/Chrome 默认用户目录，也不会关闭日常浏览器。若已有 Browser Workbench 等工具提供可用的本机 CDP 端点，可以直接连接那个端点，不需要再次启动。

CDP 只监听本机，不要开放或转发到公网。启动端口已属于其他配置时会拒绝，选择空闲端口；同一 profile 已在另一个端口运行时，应复用原端点或选择另一个专用 profile。

## 4. 让 AI 执行查询

示例请求：

> 用 customer-express-issue-query-playwright 技能，连接 http://127.0.0.1:9222，处理 D:\售后\快递问题件.xlsx。保留原表，结果另存，任务文件放在 D:\售后任务\这次查询。

AI 先读取技能说明，检查 CDP、当前账号及订单 frame，再准备 Excel、试查一批、全量续查、复查未找到并导出。只读查询不会修改聚水潭订单。

手动连接检查（按实际位置修改 `$skill`）：

```powershell
$skill = Join-Path $PWD 'skills\customer-express-issue-query-playwright'
$python = Join-Path $skill 'runtime\python\python.exe'
& $python -B "$skill\scripts\cdp_query.py" doctor --endpoint 'http://127.0.0.1:9222'
```

`ok:true` 才表示唯一订单 frame 已就绪；仍需在浏览器确认正确账号。多订单标签页时，使用返回的 `pageIndex` 指定 `--page-index N`，或只保留目标订单页面。

## 输出及续查

每个任务使用独立工作目录：`manifest.json` 是源文件与行映射；`results.json` 是已完成断点；`query-log.jsonl` 只保存计数；`summary.json` 记录导出核验；结果为 `<源文件名>_高阶版订单查询结果.xlsx`。

查询停止后，修复连接或完成登录，按原工作目录继续即可，跳过 prepare。已完成单号不会重复查询。请求失败不会记作“未找到”。源 Excel 变化、清单内部单号映射不一致或本次运行期间清单被修改会停止；清单没有跨运行防篡改签名，不应手改清单或断点。

输出追加五列：查询快递单号、发货时间、实付金额、备注、查询结果。格式化空白列及合并区域也会保留，因此结果可能在较右侧；导出摘要给出确切列范围。金额 0 保留，多订单按行对应，不合计，不覆盖原表“价值”。

新增五列宽度依次为 24、32、30、80、32，原列宽和行高保留。多订单及超长备注完整保存，原行高不足时可点选单元格查看全文。

只接受 `.xlsx`，其他格式须可靠转换为新副本。已有输出不会覆盖。账号配置、浏览器目录、源表、结果和日志都属于使用者，不应再次打进分享包。

## 验证范围

验证覆盖内置运行时启动、查询与断点单元测试、两组独立运行的 Excel 无损回归、临时 Edge 浏览器的真实 CDP 连接与本地虚构订单接口端到端测试。真实业务数据不进入测试或分享包；实际部署仍需在已登录高阶版页面试查第一批，核对字段和匹配数量。
