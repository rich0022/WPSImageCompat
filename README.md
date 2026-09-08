# WPS Image Compat for Excel

**WPSImageCompat** 是基于 TypeScript、Office.js 和 Excel JavaScript API 的跨平台 Office Add-in，目标是让 Windows / macOS Excel 显示 WPS Spreadsheet 的 DISPIMG 单元格图片。

仓库：[rich0022/WPSImageCompat](https://github.com/rich0022/WPSImageCompat)

当前版本 **0.2.0 / Phase 2**：扫描 DISPIMG 公式，并从当前工作簿中读取、解析图片资源。**尚不能显示或恢复图片**；Shape 预览属于 Phase 3。全部操作只读，不删除或替换原公式。

## 现在怎样使用

需要 Node.js 22.12+（建议受支持的 LTS）、桌面版 Excel，以及 Office.js CDN 网络访问。扫描要求 ExcelApi 1.4；图片读取单独检查 CompressedFile 1.1，宿主不支持时仍保留公式扫描结果。

```sh
npm ci
npm run certs
npm run dev
```

首次证书安装可能触发系统信任提示。开发服务固定使用 HTTPS localhost:3000。构建不会安装证书。

### macOS 加载

1. 启动服务并信任开发证书。
2. 将根目录 `manifest.xml` 复制到 `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/wps-image-compat.xml`，不存在的目录先创建。升级时替换旧清单。
3. 在 Excel 的 Insert → Add-ins → My Add-ins 中打开本插件，必要时重启 Excel。
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

**Show Images、Refresh、Remove Preview Images、Convert Workbook 均未实现，按钮保持禁用。** 三项图片设置预选开启但暂不能修改。

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
- JSZip 解包，fast-xml-parser 解析并验证 XML；不使用 Python Runtime、VBA、COM、XLL、Excel-DNA、Windows-only API 或 React。
- 同一 media 路径在一次解析中只转 Base64 一次，不持久保存工作簿或图片，不发送到服务器。
- 外部 relationship 不下载；目标必须解析到 `xl/media/`。不允许 DTD / 自定义实体声明，歧义 ID / relationship 不随机挑选。
- 限制：压缩文件 100 MiB、ZIP 条目 20,000、每 XML 4 MiB、每图片 20 MiB、累计解压读取 128 MiB。超过限制明确报错；资源流达到限制即暂停。
- Base64 读取成功不代表图片编码一定可由 Excel Shape 解码；实际图片显示格式兼容性留待 Phase 3。

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
  types/wps.ts               公共数据类型
  utils/                     XML、Base64、错误类型和文件大小限制
public/                      Ribbon 图标和本地帮助页
tests/                       单元及流程测试、合成 OOXML fixtures
manifest.xml                 XML manifest 与 Ribbon
vite.config.ts               HTTPS 开发服务和静态构建
```

解析结果使用 `Map`，重复引用共享一次解析的 Base64；跨扫描缓存、Shape renderer 和 Convert 接口留待后续阶段。

## 检查与双平台验收

```sh
npm run typecheck
npm test
npm run build
npm run validate:manifest
```

测试夹具是代码生成的 OOXML 图片部件，**不是从真实 WPS 保存的样本**。自动测试覆盖分片读取、句柄释放、ID 映射、缺失及损坏资源、路径边界和大小限制。微软 manifest 验证仅验证清单，不等于实机测试。

Windows 和 macOS 各使用真实 WPS 工作簿副本执行：

1. 验证 Ribbon 和 `Connected to Excel`。
2. 已知三张单元格图片、涉及两张表，其中一张隐藏；重复 ID 在不同位置分别计数。
3. 点击扫描，对照 XLSX 中图片部件核对 found / missing / error 和原始 ID。
4. 对普通/空白 Excel 文件得到 0；没有 WPS 部件但有公式时报告资源缺失。
5. 对不完整或损坏副本得到明确错误，不能显示旧扫描结果。
6. 连续扫描至少三次，无文件句柄占用失败；所有原公式、单元格和用户图片保持不变。
7. 对比 WPS 原文件和 Excel 导出快照，确认是否保留 cellimages 及 media；记录 Excel 版本与平台。

各阶段记录：[Phase 1](PHASE1.md)、[Phase 2](PHASE2.md)。

## GitHub 与 Cloudflare 公共使用

构建输出为 `dist/`。后续可在 Cloudflare Pages 连接本仓库，生产分支 `main`、构建命令 `npm run build`、输出目录 `dist`。

当前清单仍指向 localhost，尚未部署 Cloudflare。获得固定 HTTPS 域名后，需要生成独立生产 manifest，将其中所有 localhost URL 替换为正式域名，保留任务窗格路径 `/src/taskpane/taskpane.html`。根目录 manifest 不会自动复制到 dist，发布时需额外提供下载文件；网站根路径的安装介绍页也尚未制作。

Cloudflare 托管页面，Excel 通过 manifest 加载插件。小范围测试可侧载，企业可由 Microsoft 365 管理中心分发，公众便捷安装需提交微软加载项市场审核。正式安装后用户不需要 Node.js 或本地开发证书。

## 后续

- **Phase 3**：Shape 预览、单元格定位和缩放、去重、Refresh、仅移除插件图片。
- **Phase 4**：Convert 接口和约定范围、完善异常处理与双平台验收。

此项目与 WPS / Microsoft 没有官方关联。公开源码不自动授予开源许可；仓库许可由维护者另行选择。
