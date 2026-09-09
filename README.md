# WPS Image Compat for Excel

**WPSImageCompat** 是基于 TypeScript、Office.js 和 Excel JavaScript API 的跨平台 Office Add-in，目标是让 Windows / macOS Excel 显示 WPS Spreadsheet 的 DISPIMG 单元格图片。

仓库：[rich0022/WPSImageCompat](https://github.com/rich0022/WPSImageCompat)

当前版本 **0.5.0**：在 Phase 4 预览功能上增加简体中文/英文界面与发布更新提示。支持扫描、显示、刷新、移除预览、扫描/读取取消和诊断导出；不删除或替换原公式。Convert 已预留独立接口，实际转换尚未实现。Windows/macOS 原生 Excel 验收仍待完成，不能把单元测试视为双平台认证。

## 现在怎样使用

需要 Node.js 22.12+（建议受支持的 LTS）、桌面版 Excel，以及 Office.js CDN 网络访问。扫描要求 ExcelApi 1.4；图片读取单独检查 CompressedFile 1.1，宿主不支持时仍保留公式扫描结果。图片操作需要 ExcelApi 1.10；合并单元格整体适配需要 ExcelApi 1.13，较旧宿主按原单元格边界适配。

```sh
npm ci
npm run certs
npm run dev
```

首次证书安装可能触发系统信任提示。开发服务固定使用 HTTPS localhost:3000。构建不会安装证书。

### macOS 加载

1. 启动服务并信任开发证书。
2. 将根目录 `manifest.xml` 复制到 `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/wps-image-compat.xml`，不存在的目录先创建。升级时替换旧清单。
3. 保存当前工作并完全退出、重新打开 Excel，再打开工作簿。从 **开始（Home）→ 加载项（Add-ins）**菜单选择本插件，本项目用户已在此入口找到它。部分旧界面也可能提供“插入 → 我的加载项”旁的小箭头菜单，请以实际界面为准。侧载插件不一定出现在账户/商店对话框中；“开发工具 → Excel 加载项”管理传统加载项，不用于本项目。
4. 从 **WPS Image Compat** Ribbon 标签打开任务窗格。
5. 打开 WPS 工作簿，看到 `Connected to Excel` 后点击 **Scan Workbook**。

参考：[Microsoft macOS sideload 文档](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。

### Windows 加载

1. 在 Windows 上启动同一开发服务并信任证书。
2. 把 `manifest.xml` 放入开发共享文件夹。
3. Excel → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs，添加共享目录 UNC 地址并启用 Show in Menu。
4. 重启 Excel，在 My Add-ins → Shared Folder 中选择插件，打开任务窗格。

参考：[Microsoft Windows sideload 文档](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。菜单可能因 Excel 版本而异，组织策略可能限制侧载。

直接在浏览器访问 `https://localhost:3000/src/taskpane/taskpane.html` 只能看到界面，不能读取 Excel 工作簿。

## 语言与更新

任务窗格支持简体中文和英文，默认使用 `Office.context.displayLanguage`，连接 Excel 前使用浏览器语言；暂不支持的语言回退英文，其他中文区域回退简体中文。可在顶部手动切换，并在当前网站的本地存储中记住选择；存储受限时只在本次会话生效。切换不刷新页面，不清空扫描结果，也不修改工作簿。此功能翻译插件界面和常见错误，不翻译工作簿内容；详细宿主错误保留原文供排查。Ribbon 保持产品名称。

- 开发：`npm run dev` 使用 Vite 的开发更新能力；代码修改可能重载窗格，调试时应使用工作簿副本。
- 发布：构建同时生成带独立构建 ID 的 `version.json` 和匹配的脚本。生产窗格打开、重新获得焦点及可见状态下每 5 分钟检查同源版本信息；请求不包含工作簿内容，也不携带凭据。离线、旧部署无版本文件或检查超时不会中断图片操作。
- 新版本到达后显示“更新任务窗格”；正在扫描或修改图片时禁用此按钮。用户点击后重新加载窗格，清空本次扫描结果，已有预览保留。即使版本号相同，重新构建或回滚也可被识别；不会在图片操作中自动刷新。
- 旧版本尚无更新检测逻辑，第一次升级到此版需要关闭并重新打开窗格。开发清单始终指向本机，生产清单始终指向 Cloudflare；两者互不切换。
- 界面和逻辑更新不需要重装清单；权限、资源 URL 或 Ribbon 等 XML 清单变化仍按 Office 的更新/侧载流程处理。

详见 [更新与语言交付记录](UPDATES.md)。

## 扫描结果

| 指标 | 含义 |
| --- | --- |
| Detected | DISPIMG 单元格数量，同一个 ID 在两个单元格算两个位置 |
| Worksheets with images | 包含匹配公式的工作表数量 |
| Resources parsed | 已匹配且读取到图片二进制的单元格数量，不表示 Shape 已显示 |
| Missing resources | 无对应 ID、relationship 或 media 文件的单元格数量 |
| Resource errors | 重复 ID、非法路径、空资源等导致解析失败的单元格数量 |

公式支持 `=DISPIMG("ID_xxx",1)`、`=@_xlfn.DISPIMG("ID_xxx",1)`、可选 `@` / `_xlfn.`、大小写和空白变化。仅支持直接调用、字面量 ID 与模式 1，暂不解析嵌套函数或单元格引用参数。

扫描包含隐藏表，按批读取有值使用区域。最多展示前 100 个匹配位置和前 20 条资源问题，统计覆盖全部结果。扫描时不要编辑工作簿；不同读取阶段不是事务快照。

没有 DISPIMG 时跳过整个文件读取。普通 Excel 文件没有 `xl/cellimages.xml` 时正常返回。若宿主不支持完整文件读取，或 ZIP/XML 损坏，资源统计保持 `—`，不会将“未检查”伪装成“缺失”。

## Phase 3 图片预览

- **Scan Workbook**：只读扫描及图片资源检查。
- **Show Images**：重新扫描当前工作簿，显示找到的 PNG/JPEG 图片，已有预览跳过；问题单元格单独报告。
- **Refresh Images**：重新扫描并按当前单元格尺寸重新生成。先生成所有替换图，再删除旧预览；读取、预检或生成失败时保留旧预览并尝试清理新图。资源有缺失/错误时停止刷新。删除旧图期间的个别失败会明确报告，可能需要重试。
- **Remove Preview Images**：不依赖 WPS 源资源，遍历全部工作表并仅删除插件标记图片。重复执行不会影响普通用户图片。
- **Cancel Scan / Read**：在扫描或文件读取阶段请求取消，完成当前读取及句柄清理后结束。图片写入和清理期间禁用。
- **Download Diagnostics**：下载当前 Excel 能力、最后操作的计数和错误码，不含工作簿名称、工作表名、地址、公式、图片 ID 或图片内容，不上传数据。某些宿主可能限制下载，请按 [测试说明](TESTING.md) 记录结果。
- **Convert Workbook**：仍禁用。Phase 4 按首版范围预留 `WorkbookConverter` 接口；调用会明确返回 `CONVERSION_NOT_IMPLEMENTED`。后续转换约定输出新工作簿。

三项设置默认开启。Keep aspect ratio / Fit image inside cell 可修改；设置变化通过 Refresh 应用到旧预览。Preserve original DISPIMG formula 在预览模式始终开启且不可取消。

默认以 Excel 的点为单位计算尺寸，保持比例、居中、预留最多 1 点内边距。取消适配后保留图片原始 Shape 尺寸，可能超出单元格；靠近工作表左上角时位置钳制到 0。取消比例但保持适配会拉伸填充单元格。隐藏行/列或零尺寸单元格跳过；取消隐藏后可重新 Show 或 Refresh。隐藏工作表仍处理。

Shape 使用 Excel 原生 `placement = TwoCell`，跟随单元格移动和缩放，比例设置同时写入 `lockAspectRatio`。不同 Excel 版本在非等比调整行高/列宽、合并和隐藏单元格时可能有不同效果；没有后台事件监听器持续重新居中，必要时点击 Refresh 重新适配。

名称格式为 `WPSIMG_<ID 的 SHA-256 摘要>_<A1 地址>_<运行 UUID>`，完整图片 ID 放在替代文字描述中。每次运行有唯一名称；Show 根据 ID 和当前位置判断已有图片，避免同一 ID 在多个单元格或插入行后错误去重。

删除需要名称以 `WPSIMG_` 开头 **并且**替代文字标题为 `WPS Image Compat preview v1`。普通图片仅名字相似不会被删除。手动改掉标记或分组后的预览不保证被识别，请保留这些标记；清理只遍历顶层 Shape，不递归解组用户对象。

**这些是工作簿中的真实浮动 Shape。保存文件会保存预览；它们不是 Excel 原生单元格图片，不会自动消失。** 原公式一直保留，Remove 后原有公式错误显示可能重新出现。图片解析在本机完成，无上传。Excel 的 WPS 资源保留问题仍适用。

## Phase 2 工作流程

```text
Excel 全部工作表 → DISPIMG 单元格和 ID
Office getFileAsync(Compressed) → 分片读取 → 完整 XLSX 二进制 → 关闭文件句柄
JSZip → xl/cellimages.xml 中 cNvPr.name → blip embed relationship ID
      → xl/_rels/cellimages.xml.rels 中 Target → xl/media/* → Base64
ID 与单元格匹配 → found / missing / error
```

- parser 不依赖 Office.js；所有 WPS 特有部件路径和 XML 结构解释集中在 `wps-image-parser.ts`。
- `workbook-reader.ts` 只获取二进制，成功和失败路径都会尝试关闭文件句柄。
- `getFileAsync`、`getSliceAsync`、`closeAsync` 每次回调等待上限 30 秒，不自动重试；取消/超时后迟到的文件句柄仍尝试关闭。Excel 批处理的 `context.sync()` 不能被中断，扫描取消在批次边界生效；解包取消在当前解析结束后生效，不承诺即时停止。
- JSZip 解包，fast-xml-parser 解析并验证 XML；不使用 Python Runtime、VBA、COM、XLL、Excel-DNA、Windows-only API 或 React。
- 同一 media 路径在一次解析中只转 Base64 一次，不持久保存工作簿或图片，不发送到服务器。
- 外部 relationship 不下载；目标必须解析到 `xl/media/`。不允许 DTD / 自定义实体声明，歧义 ID / relationship 不随机挑选。
- 限制：压缩文件 100 MiB、ZIP 条目 20,000、每 XML 4 MiB、每图片 20 MiB、累计解压读取 128 MiB。超过限制明确报错；资源流达到限制即暂停。
- Base64 读取成功不代表图片编码一定可由 Excel Shape 解码；Phase 3 仅传入官方支持的 PNG/JPEG，其他格式报告跳过，损坏编码由宿主报错。

[微软 getFileAsync 文档](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js)列出桌面 Excel Windows/macOS 的 Compressed 支持，但具体版本和 WPS 自定义部件能否在 Excel 导出的快照中保留，仍需实机验证。如果 Excel 已丢弃图片部件，本阶段不会从不存在的资源恢复图片，也不自动读取磁盘原文件。

## 项目结构

```text
src/
  taskpane/                  任务窗格界面、宿主初始化和统计展示
  core/
    dispimg-formula.ts        纯公式识别和 A1 坐标转换
    dispimg-scanner.ts        Excel 只读扫描
    workbook-reader.ts        Compressed 文件读取与资源释放
    wps-image-parser.ts       WPS OOXML、relationship、media 解析
    image-mapper.ts           单元格和图片资源匹配
    workbook-detector.ts      扫描、读取、解析的流程协调
    image-renderer.ts         Shape 插入、预检、回滚和清理
    image-layout.ts           纯几何布局计算
    preview-identity.ts       唯一名称和所有权/位置判断
    preview-controller.ts     Show/Refresh 前重新扫描
    workbook-converter.ts     独立转换接口与未实现状态
    diagnostics.ts            不含工作簿内容的诊断报告
  types/wps.ts               公共数据类型
  utils/                     XML、Base64、错误、取消、超时和文件大小限制
scripts/diagnose-workbook.ts  本地只读真实 XLSX 诊断工具
public/                      Ribbon 图标和本地帮助页
tests/                       单元及流程测试、合成 OOXML fixtures
manifest.xml                 XML manifest 与 Ribbon
vite.config.ts               HTTPS 开发服务和静态构建
```

解析结果使用 `Map`，重复引用共享一次解析的 Base64；不保留跨扫描资源缓存，Show/Refresh 每次读取最新快照。

## 检查与双平台验收

```sh
npm run typecheck
npm test
npm run build
npm run validate:manifest
npm run validate:manifest:production
npm run diagnose:file -- "/absolute/path/sample.xlsx" "/absolute/path/new-report.json"
```

自动测试夹具是代码生成的 OOXML 图片部件，**不是从真实 WPS 保存的样本**。自动测试覆盖分片读取、句柄释放、ID 映射、缺失及损坏资源、路径边界、大小限制、取消/超时、诊断隐私及预览回滚。微软 manifest 验证仅验证清单，不等于实机测试。

本地诊断工具直接用项目 parser 读取磁盘文件，统计 worksheet XML 中的直接 DISPIMG 调用，并核对读前/读后的 SHA-256；不会重存 XLSX。报告采用只创建新文件的方式，拒绝覆盖输入或已有报告。它不经过 Office.js，也不证明 Excel 导出的快照保留了 WPS 部件；共享公式的省略公式单元格不在该辅助工具的统计范围内。真实工作簿和诊断报告应留在仓库外，禁止提交客户数据。完整步骤和结果记录表见 [TESTING.md](TESTING.md)。

Windows 和 macOS 各使用真实 WPS 工作簿副本执行：

1. 验证 Ribbon 和 `Connected to Excel`。
2. 已知三张单元格图片、涉及两张表，其中一张隐藏；重复 ID 在不同位置分别计数。
3. 点击扫描，对照 XLSX 中图片部件核对 found / missing / error 和原始 ID。
4. 对普通/空白 Excel 文件得到 0；没有 WPS 部件但有公式时报告资源缺失。
5. 对不完整或损坏副本得到明确错误，不能显示旧扫描结果。
6. 连续扫描至少三次，无文件句柄占用失败；所有原公式、单元格和用户图片保持不变。
7. 对比 WPS 原文件和 Excel 导出快照，确认是否保留 cellimages 及 media；记录 Excel 版本与平台。

Phase 3 还需在 Windows/macOS 各执行以下真实工作簿验收：

- 放置横图、竖图、重复 ID、不连续位置和用户自有图片；Show 后检查居中、不越界和原始比例。
- 重复 Show 不新增重复图；两个单元格引用相同 ID 时各显示一张。
- 插入/删除行、改变行高/列宽、合并单元格后检查原生附着行为，并执行 Refresh 验证重新适配。
- 修改两项设置后 Refresh，核对比例和填充差异；取消适配时检查大图重复执行不累加。
- 保护工作表、损坏图片或移除资源后尝试 Refresh，核对旧预览保留和失败提示。
- Remove 后仅插件图片消失，原公式、用户图片及其他数据保持不变。
- 保存、关闭、重新打开，核对预览和所有权标记保留，仍能再次 Show 去重及 Remove。

本轮自动测试使用 Excel 模拟对象，不能证明真实 Excel 布局、原生锚点或保存行为一致。

各阶段记录：[Phase 1](PHASE1.md)、[Phase 2](PHASE2.md)、[Phase 3](PHASE3.md)、[Phase 4](PHASE4.md)。

## GitHub 与 Cloudflare 公共使用

本项目已配置 **Cloudflare Workers Static Assets**，Worker 名为 `wpsimagecompat`，生产域名为 `https://wpsimagecompat.fogce.workers.dev`。这不是 Pages 项目。

Cloudflare 的 Worker → Settings → Build 设置：生产分支 `main`，构建命令 `npm run build`，部署命令 `npx wrangler deploy`，根目录为仓库根目录。配置来自仓库的 `wrangler.jsonc`，直接部署 `dist/`，不会进入 Vite 自动改写流程，也不需要 Cloudflare Vite 插件。

构建结束会生成 `dist/manifest.xml`，其中全部 URL 指向生产域名；根目录 `manifest.xml` 继续用于 localhost 开发。网站首页提供安装说明和生产清单下载。安装生产版时替换旧开发版清单，两者使用相同加载项 ID。

本地验证可运行 `npm run deploy:check`，实际发布可运行 `npm run deploy`（需要对应 Cloudflare 账号权限）。部署步骤、验收与排错见 [DEPLOYMENT.md](DEPLOYMENT.md)。配置和 dry run 通过不等于生产部署成功，应核对线上入口。

Cloudflare 托管页面，Excel 通过 manifest 加载插件。小范围测试可侧载，企业可由 Microsoft 365 管理中心分发，公众便捷安装需提交微软加载项市场审核。正式安装后用户不需要 Node.js 或本地开发证书。

## 后续

- 完成 Windows/macOS 原生 Excel 验收，记录实际版本、能力、布局和保存结果。
- 单独设计并实现输出新工作簿的实际 Convert，不改变预览操作的只保留原公式约定。

此项目与 WPS / Microsoft 没有官方关联。公开源码不自动授予开源许可；仓库许可由维护者另行选择。
