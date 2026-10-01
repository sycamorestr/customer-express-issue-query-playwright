# 客服快递问题件查询 · Playwright + CDP

读取快递问题件 Excel，在聚水潭**高阶版**订单页按快递单号批量查询发货时间、实付金额和备注，将结果追加到新 Excel，保留源表内容、公式、格式和附件。

使用 Playwright 连接 Edge/Chrome 的 CDP 端口，无需 OpenCLI 或浏览器扩展。查询仅调用只读订单接口，不修改聚水潭订单。

## 下载与使用

**普通使用者请下载 [最新 Release 的 Windows64 便携完整包](https://github.com/sycamorestr/customer-express-issue-query-playwright/releases/latest)**。便携包内已带 Python、Playwright、Excel 依赖、安装器和自检，不需要自行安装 Python/npm 或下载 Playwright 浏览器。

1. 完整解压 Release 附件中的便携 ZIP，运行 `check.cmd`。
2. 运行 `start-browser.cmd`，在专用 Edge/Chrome 浏览器登录自己的聚水潭高阶版，打开订单列表。
3. 让 AI 读取并使用 `skills/customer-express-issue-query-playwright/SKILL.md`，提供 Excel 路径和本次任务工作目录。

首次默认连接地址为 `http://127.0.0.1:9222`。已有 Browser Workbench 等工具提供的本机 CDP 端点也可直接复用。首次登录、验证码由使用者本人完成；登录态保存在其电脑的独立浏览器目录。

详细安装、参数和续查步骤见 [便携包使用说明](PORTABLE.md)；完整业务规则见 [SKILL.md](skills/customer-express-issue-query-playwright/SKILL.md)。

GitHub 的 **Code → Download ZIP** 和自动生成的 **Source code** 是源码，不含运行时；需要开箱使用时，应下载 Release 中名称含“便携完整包”的附件。`check.cmd`、`install.ps1` 和 `batch.ps1` 使用内置运行时，在完整便携包中运行。

## 功能

- 支持多个工作表、同格多个单号、跨行重复单号和一个单号多个订单。
- 精确表头优先识别“月结赔付单号”，其次“快递单号/物流单号/运单号/发货单号”，最后以“单号”兜底；同级候选歧义时停止。
- 查询键自动忽略首尾多余空白、逗号、顿号、分号和竖线；保留源单元格，未知说明及不完整单号仍需核对。
- 按 `l_id`、`@=` 查询，同时提交原单号与 `@` 前缀形式；默认不加日期限制。
- 每批默认 25 个单号；返回达到 500 条时自动拆批，单号仍达到上限则停止，防止漏回填。
- 每批校验后保存断点；中断后沿用原工作目录续查，已完成项自动跳过；支持小批复查未找到。
- 连接、登录、权限、超时或接口异常会停止，不误报“未找到订单”。
- 直接追加 XLSX 工作表 XML，逐项核验原单元格、公式、格式、图片及附件部件；不重存原工作簿。

追加列依次为“查询快递单号”“发货时间(confirm_date)”“实付金额(paid_amount)”“备注(remark)”“查询结果”。金额 0 保留，多订单逐行对应，不合计、不覆盖原表“价值”。只直接支持 `.xlsx`。

新增五列宽度依次为 24、32、30、80、32；原列宽、默认列宽和原行高保持不变。多订单及长备注完整存储，原行高不足时可点选单元格查看全文。

## 源码开发与验证

源码仓库保留通用脚本和测试，便携运行时作为 Release 附件分发。Windows x64 开发环境可使用 Python 3.13、Node.js 24，并安装固定版本依赖：

```powershell
python -m pip install -r skills/customer-express-issue-query-playwright/scripts/requirements.txt
python -B -m unittest discover -s tests -p "test_*.py" -v
python -B -m unittest discover -s skills/customer-express-issue-query-playwright/tests -p "test_*.py" -v
node --test tests/test_query.js
```

两组 Python 测试分别在独立进程运行，避免同名测试模块冲突；技能目录内的回归测试也会随技能安装，便于独立维护时检查表头、单号解析和无损导出。

源码环境直接调用 `python skills/customer-express-issue-query-playwright/scripts/cdp_query.py doctor --endpoint http://127.0.0.1:9222`；准备和导出使用同目录的 `excel.py`。专用浏览器仍可通过 `start-browser.cmd` 启动。具体查询顺序和参数以技能说明为准。

主要文件：

| 路径 | 用途 |
|---|---|
| `skills/customer-express-issue-query-playwright/SKILL.md` | 查询业务规则与操作流程 |
| `scripts/cdp_query.py`（技能目录内） | CDP 连接、订单 frame 选择、锁和断点 |
| `scripts/query.js`（技能目录内） | 登录页面中的异步只读查询与拆批 |
| `scripts/excel.py`（技能目录内） | 表头识别、单号解析及无损 Excel 导出 |
| `scripts/start_browser.ps1`（技能目录内） | 独立 Edge/Chrome 配置及回环调试端口 |
| `tests/` | 查询、断点、Excel 离线测试和可选本地 CDP 集成测试 |
| `tests/`（技能目录内） | 表头兜底、边界分隔符、列宽及无损保留回归测试 |

验证覆盖离线回归、临时 Edge 的真实 CDP 连接、本地虚构订单接口的完整查询/续查/复查/导出，以及中文路径解压与安装。真实业务数据不进入测试或仓库；实际使用时仍须先在本人已登录的高阶版页面试查一批，确认账号、字段和匹配数。

`tests/smoke_cdp.py` 是可选集成检查，使用专用临时浏览器及本机虚构接口，不访问真实 ERP。需要 Windows 和 Edge；使用时传入新的空目录：

```powershell
python -B tests/smoke_cdp.py --workdir D:\临时测试\cdp-smoke
```

## 依赖与数据

依赖版本、下载地址和 SHA256 记录在 [python-dependencies.json](python-dependencies.json)、[download-provenance.json](download-provenance.json)，许可证说明见 [THIRD_PARTY.md](THIRD_PARTY.md)。完整便携包内部带文件校验清单，Release 另附 ZIP 的 SHA256 文件。

仓库和便携包不含客户 Excel、订单结果、浏览器配置、Cookie 或登录凭据。实际业务工作目录和浏览器登录目录应放在仓库之外，CDP 仅供本机使用。
