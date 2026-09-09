# 0.6.0 — 只读兼容检查

## 本版变化

新增“检查兼容性”和“下载详细兼容报告”，保留中英切换、发布更新提示以及现有图片预览流程。兼容扫描不调用写入、重新计算、断开链接或转换接口。

- 公式规则区分宿主返回的公式错误与函数兼容标记风险；记录 DISPIMG 和显式失效引用，不将所有错误归因于 WPS。
- 显式外部公式引用与 workbook.xml 外部引用 ID 只做盘点；不解析外部目标关系、不访问网络或目标文件，不能据此判定链接可用/失效。
- 日期系统报告 1900/1904，合法系统不判错；不猜测日期、不修改单元格值。
- 分项列出判断、完整/部分/不可用状态、错误码、单元格位置及实际证据。当前单元格与 OOXML 快照分别标明来源。
- 取消及失败保留“未检查”边界。扫描最多 100 万个单元格、每批最多 5,000 格；保留 5,000 条详情，页面显示前 100 条。省略数量明确列出，不因详情截断而宣称某类别无问题。
- 详细 JSON 包含业务公式和引用，只在用户点击后本地下载。简要诊断内容不扩大，文件与内容均不自动上传。

## 修改文件清单

| 文件 | 作用 |
| --- | --- |
| `src/core/compatibility/types.ts` | 统一检查、发现和报告结构及详情上限 |
| `src/core/compatibility/formula-rules.ts` | 独立、可测试的公式规则 |
| `src/core/compatibility/scanner.ts` | 只读 Excel 批量扫描与完成状态 |
| `src/core/compatibility/metadata-parser.ts` | 与 Office 解耦的 ZIP/XML 元数据解析 |
| `src/core/compatibility/controller.ts` | 汇总实时单元格与文件快照结果 |
| `src/taskpane/compatibility-view.ts` | 本地化报告展示，证据按纯文本渲染 |
| `src/taskpane/taskpane.ts` | 检查操作、取消/忙碌保护、本地下载 |
| `src/taskpane/taskpane.html`、`taskpane.css`、`i18n.ts` | 新入口、报告布局和中英文本 |
| `tests/compatibility.test.ts` | 公式、误报、元数据、损坏/限制/取消、只读宿主和汇总测试 |
| `package.json`、`package-lock.json` | 仅将项目版本升级至 0.6.0；无新依赖 |
| `public/index.html`、`public/support.html` | 发布入口和帮助说明 |
| `README.md`、`TESTING.md` | 使用范围和双平台验收说明 |
| `COMPATIBILITY_ROADMAP.md`、`RELEASE_0_6.md` | 已实现边界、后续方向及交付记录 |

原有 DISPIMG parser、图片渲染器、读取模块、manifest ID/权限/版本、Cloudflare 配置保持现状。本版无需因界面更新重装 manifest；本地开发清单与生产清单仍分别使用原地址。

## 验证结果

- `npm run typecheck`：通过。
- `npm test`：81 项通过，无失败。
- `npm run build`：通过，生成匹配的版本信息、脚本及生产清单。
- 本地浏览器：0.6.0 页面显示新增入口；中英切换通过；非 Excel 宿主下检查、下载和图片操作均禁用。
- 用户授权样本：直接从磁盘使用新元数据 parser 读取，结果为 1900 日期系统、0 个外部引用声明；读前/读后 SHA-256 一致。未重存文件，样本、文件路径与内容未提交。

以上磁盘测试不能证明 Excel 的文件快照与原文件完全相同，浏览器界面验证也不能代替真实 Excel 会话。

## 仍存在的问题与限制

- Windows/macOS 原生 Excel 内的新扫描、详细报告下载及已有图片流程仍需按 TESTING.md 实测；本轮无 Windows 宿主。
- 不检查定义名称、INDIRECT 文本目标、外部目标关系/可达性、连接、图表、宏、格式布局或日期值语义。
- 单元格逐批读取与文件快照读取不是原子操作，扫描期间编辑可能导致证据时点不同。
- 元数据输入上限 100 MiB、ZIP 20,000 项、workbook.xml 解压上限 4 MiB；超限明确未检查，不做无限解压。
- 外部引用记录不是目标可用性验证；普通工作簿“检查范围内未发现问题”也不是全格式兼容认证。
- 正式 Convert / 自动修复尚未实现。插件宿主仍为 Microsoft Excel，不承诺直接在 WPS 中安装运行。
