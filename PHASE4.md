# Phase 4 交付记录

版本：应用 0.4.0，XML manifest 1.3.0.0。

## 范围与结果

按首版约定预留 Convert 独立接口；实际转换不在本轮实现，按钮保持禁用，调用接口明确拒绝并返回 `CONVERSION_NOT_IMPLEMENTED`。后续接口约定输出新工作簿。预览操作始终保留原公式。

本轮完成扫描/读取取消、文件 API 回调超时与迟到句柄清理、隐私诊断导出、本地只读 XLSX 检查工具，以及自动测试和双平台验收说明。没有新增依赖、改用平台专有技术或引入 UI 框架。

## 修改文件

- `src/core/workbook-converter.ts`：转换接口与明确未实现状态。
- `src/core/diagnostics.ts`：宿主能力、计数、问题码；不导出工作簿内容。
- `src/utils/office-async.ts`、`src/utils/cancellation.ts`：30 秒文件回调超时、取消及迟到成功清理。
- `src/core/workbook-reader.ts`：将超时/取消接入文件打开与分片读取，保证已获取句柄尝试关闭。
- `src/core/dispimg-scanner.ts`、`src/core/workbook-detector.ts`：批次边界取消；取消不伪装成缺失资源。
- `src/core/preview-controller.ts`：读取阶段可取消，进入 Shape 写入前关闭取消入口。
- `src/taskpane/taskpane.ts`、`src/taskpane/taskpane.html`：取消、诊断下载和转换状态说明。
- `scripts/diagnose-workbook.ts`：复用 parser 读取磁盘 XLSX，核对原文件散列，报告输出禁止覆盖已有文件。
- `tests/office-async.test.ts`、`tests/diagnostics.test.ts`、`tests/diagnose-file.test.ts`：新增回调、隐私、转换边界及 CLI 原件保护测试。
- `tests/workbook-reader.test.ts`、`tests/workbook-detector.test.ts`、`tests/preview-controller.test.ts`、`tests/dispimg-scanner.test.ts`：补充取消与未写入验证。
- `package.json`、`package-lock.json`、`manifest.xml`、`tsconfig.json`：版本、诊断命令与工具脚本类型检查。
- `README.md`、`TESTING.md`、`PHASE4.md`、`public/index.html`：使用方法、双平台测试流程和交付边界；纠正 Mac 侧载下拉菜单与账户/商店对话框混淆的说明。

## 验证

- TypeScript：通过。
- 自动测试：61/61 通过，覆盖原有解析/显示流程及本轮异常路径。
- 真实样本：磁盘 OOXML 解析、全部图片像素解码和布局公式检查完成，原始文件散列未变。样本、图片、路径及详细报告留在仓库之外。
- Vite 构建：通过。开发和生产 XML 清单均通过微软官方验证服务；验证只证明清单有效，不证明原生 Excel 运行结果。
- 本机已安装 Excel，但没有可控制的 Excel 会话；原生任务窗格、getFileAsync、Shape、保存行为及下载流程未测。Windows 原生测试未执行。

## 仍存在的问题

1. 实际 Convert 尚未实现；没有把预览重命名成转换结果。
2. Excel 打开/保存后是否保留 WPS 自定义部件仍取决于真实宿主；磁盘解析成功不等于 Excel 快照可读。
3. 取消不是中断 Excel 批处理或 JSZip 当前解析；当前工作结束后才生效。文件关闭若宿主不回调，最多再等待 30 秒并报告/保留原错误。
4. 原生锚点、合并/隐藏单元格、保存再打开、诊断下载需 Windows/macOS 分别按 TESTING.md 验收。
5. 诊断 CLI 不展开 worksheet XML 中省略的共享公式；只作直接调用的磁盘基线。
6. 沿用既有开发依赖风险记录，见 DEPLOYMENT.md；本轮不做无关依赖升级。

本轮本地测试不代表 Cloudflare 线上版本已经更新。
