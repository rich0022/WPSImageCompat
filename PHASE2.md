# Phase 2 交付记录

日期：2026-09-09。版本 0.2.0，manifest 1.1.0.0。仅进入 Phase 2，未实现 Phase 3 Shape 或 Phase 4 转换。

## 已实现

- `getFileAsync(Compressed)` 宿主能力检查、1 MiB 顺序分片读取、长度/顺序校验、成功/失败路径释放文件句柄。
- JSZip 解包；fast-xml-parser 验证和解析 WPS cellimages、relationship 与 Content Types XML。
- DISPIMG ID → relationship ID → xl/media → MIME / Base64 映射。
- 同一 media 在一次解析中复用 Base64，重复图片 ID 或 relationship ID 明确报错。
- 普通 Excel 缺少 cellimages 正常返回，资源缺失与解析失败分别统计。
- Scan Workbook 完成全部阶段；文件读取失败仍保留公式数量，资源统计显示未检查。
- 禁止外部资源下载、DTD、越界 media 路径；对输入及流式解压设置大小限制。
- 补充 Windows/macOS 安装、升级和真实文件验收步骤。用户数据与公式完全只读。

## 验证

| 检查 | 结果 |
| --- | --- |
| TypeScript strict | 通过 |
| 单元及流程测试 | 26 项通过 |
| Vite 生产构建 | 通过 |
| 微软 manifest 在线验证 | 通过 |
| npm 依赖安全审计 | 0 vulnerabilities |
| 真实 WPS 文件样本 | 本轮未提供；测试使用合成 OOXML 图片部件 |
| Windows/macOS Excel 完整流程 | 待实机验收 |

测试包括 Office callback 适配、分片合并、出错关闭文件、普通 ZIP、ID 到确切 Base64 的映射、命名空间前缀变化、UTF-16 XML、关系及图片缺失、重复 ID、非法路径、损坏 XML、外部 relationship、空图片、解压上限，以及流程协调和失败统计。

## 本阶段新增文件

- `src/core/workbook-reader.ts`
- `src/core/wps-image-parser.ts`
- `src/core/image-mapper.ts`
- `src/core/workbook-detector.ts`
- `src/utils/base64.ts`
- `src/utils/xml.ts`
- `src/utils/errors.ts`
- `src/utils/limits.ts`
- `tests/fixtures/wps-workbook.ts`
- `tests/workbook-reader.test.ts`
- `tests/wps-image-parser.test.ts`
- `tests/workbook-detector.test.ts`
- `PHASE2.md`

## 本阶段修改文件

- `src/types/wps.ts`：解析结果和诊断类型。
- `src/taskpane/taskpane.ts`：接入流程协调器及诊断显示。
- `src/taskpane/taskpane.html`：Phase 2 标识、资源统计。
- `manifest.xml`：更新版本和功能描述。
- `package.json` / `package-lock.json`：0.2.0、GitHub 地址、fast-xml-parser 依赖。
- `.gitignore`：排除本地凭据及真实工作簿。
- `public/support.html`：更新帮助和仓库链接。
- `README.md`：当前使用方法、行为、限制和双平台测试说明。

## 仍存在的问题与边界

1. Windows/macOS 上 CompressedFile 支持以及 Excel 是否保留 WPS 自定义部件，需要真实文件验收；读取快照缺少部件时只能报告缺失。
2. “Resources parsed” 表示二进制已读取，不代表已解码验证所有图片格式或插入 Shape。
3. 不支持嵌套 DISPIMG、引用参数或其他显示模式；沿用 Phase 1 公式范围。
4. 不支持加密工作簿、外部图片、超过配置上限的文件；ZIP 条目内容不做全包 CRC 扫描。
5. 扫描与导出不是原子快照，扫描期间应避免编辑；极稀疏、跨度大的工作表可能较慢。
6. Office callback 如果永久不返回，当前操作会持续等待；当前没有自动超时或取消。
7. JSZip 3.10.1 的 internalStream 存在于运行时但未在 JSZipObject 类型中声明，局部类型补充用于限制解压读取。升级 JSZip 时应运行流式上限测试。
8. Shape 显示、刷新、移除、Convert 仍禁用；尚未部署到 Cloudflare。

发布范围：首次 Git 提交包含 Phase 1 和 Phase 2 的项目源码、测试、图标和文档。node_modules、dist、真实 Excel 文件、开发证书、本地侧载副本和环境凭据不提交。
