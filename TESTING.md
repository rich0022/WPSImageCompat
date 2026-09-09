# Windows / macOS 测试说明

先使用原文件的独立副本；不要直接在客户文件或唯一原件上测试 Shape、刷新、保存。测试样本和报告留在仓库之外。插件不自动上传工作簿。

## 本地准备

```sh
npm ci
npm run certs
npm run dev
```

开发服务为 HTTPS localhost:3000。macOS 把根目录 `manifest.xml` 复制到 `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/wps-image-compat.xml`，保存工作并完全退出、重开 Excel。从“开始 → 加载项”菜单选择插件（本项目用户已确认此入口可见）。部分旧界面有“插入 → 我的加载项”旁的小箭头，请以实机为准。不要用账户/商店的空列表判断侧载失败，“开发工具 → Excel 加载项”也不是本项目的入口。Windows 按 README 配置受信任共享文件夹。开发清单和生产清单使用同一 ID，测试时只保留所需版本。窗格应显示版本 `0.5.0` 和语言选择；若仍是旧版，关闭并重新打开窗格。

参考：[微软 macOS 侧载步骤](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)、[Windows 共享目录侧载](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。

浏览器直接打开任务窗格只能验收页面和宿主缺失提示，不能执行真实 Excel API。

## 先检查磁盘样本

```sh
npm run diagnose:file -- "/absolute/path/sample.xlsx" "/absolute/path/new-report.json"
```

工具只读原始 ZIP 和 XML，复用正式 WPS parser；报告计数、资源问题码和输入散列，不输出图片或公式。已有报告不覆盖。它读取的是磁盘文件，不能反映 Excel 尚未保存的编辑。worksheet XML 中共享公式的省略单元格不展开，因此它只作为直接 DISPIMG 调用的基线；真正的 Scan 使用 Excel 返回的单元格公式。

## 在两个平台分别执行

开始前记录 Excel 版本、系统版本、工作簿副本名（仅私人记录）、原始公式和用户图片数量。打开插件确认 `Connected to Excel`。

| 操作 | 预期结果 |
| --- | --- |
| Scan Workbook | DISPIMG 数量与已知基线一致，资源 found/missing/error 可解释，公式不变 |
| 无 DISPIMG 的普通空白文件 | 0，不进行文件读取，无错误 |
| 有 DISPIMG，但没有 WPS 图片部件 | 明确缺失资源；不误报成功 |
| Show Images | 可见目标单元格显示 PNG/JPEG，默认保持比例、居中、不越界，公式不变 |
| 再次 Show | 已有图片计入 already shown；不产生重复图 |
| 相同 ID 位于两个单元格 | 两个位置分别显示图片 |
| 改行高/列宽、插入行、合并区域后 Refresh | 按当前目标范围重新适配；用户自有图片不受影响 |
| 保护工作表、资源缺失或图片损坏后 Refresh | 明确报告失败；预检/生成失败保留旧图，必要时说明回滚问题 |
| Remove Preview Images 两次 | 只删除带插件名称前缀且所有权标记正确的图片；第二次移除 0 |
| 连续 Scan 三次 | 无文件句柄占用错误，结果一致 |
| Scan 或 Show 读取时 Cancel | 等待当前批次/解析及句柄清理后显示取消；不开始图片写入；可再次操作 |
| 图片写入或 Remove 期间 | 取消按钮禁用，等待操作结束 |
| Download Diagnostics | 得到 JSON，检查 host 能力、计数、错误码；没有公式/图片/单元格等内容 |
| Convert Workbook | 禁用并说明尚未实现；不更改任何内容 |
| 语言切换 | 自动跟随 Excel；手动中英切换立即生效，扫描计数/位置和预览设置不变；重开后保留选择 |
| 发布更新提示 | 发布不同构建 ID 后打开/聚焦窗格，出现更新按钮；无更新时不显示 |
| 图片操作期间遇到更新 | 提示可见但更新按钮禁用，完成后方可点击；不能打断 Shape 写入 |
| 点击更新 | 窗格重新加载到新版，扫描结果清空，工作簿公式和已有预览保留 |
| 检查更新时断网 / 404 / 错误 JSON | 不影响扫描/图片流程，不显示错误更新通知 |
| 保存副本、关闭并重开，Show / Remove | 所有权标记仍保留，去重和清理继续有效；原公式不变 |

还要检查隐藏表、隐藏行/列、横图/竖图，以及取消 Keep aspect ratio / Fit image inside cell 后的表现。TwoCell 是原生附着方式，行列变动后不保证持续重新居中，Refresh 才重新计算布局。隐藏的零尺寸目标可以跳过，不能要求所有引用都新增 Shape。

## 常见结果的判断

- `UNSUPPORTED_HOST`：资源统计未检查；更新桌面 Excel 并用诊断报告核对 CompressedFile 能力。[微软支持表](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js)列出桌面 Windows/macOS Excel 支持 Compressed，运行时仍以宿主能力检查为准。
- 磁盘诊断找到资源，窗格却找不到：Excel 的 `getFileAsync` 快照可能没有保留 WPS 自定义部件。不要把磁盘结果当作快照证据；记录版本和窗格报告。本版没有文件选择回退。
- `OFFICE_TIMEOUT`：单次文件回调等待超过 30 秒。关闭并重新打开窗格；应用不自动重试。迟到的文件句柄会尝试清理，宿主完全不响应时无法保证释放成功。
- `OPERATION_CANCELLED`：扫描/读取停止，不能推导此前就已存在的预览是否被移除。
- Refresh 报告部分删除失败：不要反复保存后覆盖原件；检查副本中旧图/新图与工作表保护，按提示重试。
- 下载被 WebView 拦截：记录宿主版本和提示。本版诊断下载还需真实宿主验收。

## 验收记录

| 层级 | 本轮状态 | 能证明什么 |
| --- | --- | --- |
| TypeScript / 构建 / 自动测试 | 见 PHASE4.md | 类型、打包、纯逻辑和模拟 Office 流程 |
| 用户授权真实样本的磁盘读取 | 已完成，报告仅留本地 | 原始 OOXML、ID 映射和原文件散列完整性 |
| 真实样本图片解码与布局公式 | 已完成，报告仅留本地 | 图片像素可解码及几何计算，不含 Excel 显示 |
| macOS 原生 Excel UI / getFileAsync / Shape | 待实机验收 | 不能用本地文件读取替代 |
| Windows 原生 Excel UI / getFileAsync / Shape | 待实机验收 | 本轮没有 Windows 宿主 |

每个平台完成后记录日期、Excel/系统版本、每项结果、失败错误码和诊断 JSON。没有实际执行的项目写“未测”，不能写“通过”。

## 0.6 只读兼容报告：双平台追加验收

在 Windows 和 macOS 各用新建测试副本执行，并记录版本与报告；不要覆盖业务原件。

| 场景 | 预期 |
| --- | --- |
| 普通文件，`=SUM(A1:A2)`、合法 `@`、表格结构引用 | 不因语法外观产生兼容故障；日期系统合法 |
| `=1/0` 与文本 `#DIV/0!` | 只把宿主实际公式错误报告为问题；文本不误判 |
| DISPIMG、`_xlfn` 公式 | 兼容标记为风险，实际错误值独立记录；保留公式 |
| 显式外部工作簿引用、`#REF!` | 前者是依赖风险且目标未检查，后者记录已观察到的失效引用 |
| 1900 与 1904 文件 | 如实盘点，不标故障，不改日期序号 |
| CompressedFile 不可用或文件读取失败 | 公式结果保留，依赖快照的项目明确未完成 |
| 扫描取消、宿主读取失败或超出 100 万格 | 不显示全部通过；可再次扫描，无数据写入 |
| 超过 5,000 条发现 | 实际计数保留，界面和下载说明详情被截断 |
| 切换中英文 | 报告结论、证据和数量保留，标题和建议正确翻译 |
| 下载详细兼容报告 | 本地 JSON 含位置/公式/证据/范围；无网络上传；核对敏感信息后才分享 |
| 下载简要诊断 | 仍无公式、地址、文件内容，不因兼容检查扩大隐私范围 |
| 扫描中发布新版 | 更新按钮禁用；完成或取消后才允许重载 |
| 执行其他操作或重载窗格 | 旧兼容报告清空，不能下载失效结果 |

本轮自动化与磁盘结果见 RELEASE_0_6.md。上述原生宿主场景目前仍标记为“待测”，不能由浏览器页面或模拟对象测试推导通过。
