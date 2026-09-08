# Phase 3 交付记录

日期：2026-09-09。版本 0.3.0，manifest 1.2.0.0。

## 完成的功能

- `worksheet.shapes.addImage(base64)` 插入 PNG/JPEG 预览。
- 用 Excel Shape 原始宽高计算比例，按单元格点坐标居中并适配；支持可选拉伸或保持原始大小。
- 设置 `placement = TwoCell` 与 `lockAspectRatio`，使用原生单元格附着行为。
- 可用时适配整个合并区域（ExcelApi 1.13），否则按锚单元格；隐藏行列跳过，隐藏工作表照常处理。
- 唯一 WPSIMG 名称、所有权标记和完整图片 ID 元数据；按当前图片位置去重，兼顾相同 ID 多位置和原始尺寸大图。
- Show / Refresh 每次重新扫描当前工作簿，不复用上次扫描快照。
- Refresh 先生成新预览，全部成功后移除旧图。生成失败尝试回滚新图，保留旧图；失败清理和删除错误明确报告。
- Remove 只移除具有保留前缀及所有权标题的顶层 Shape，不删除仅名称相似的用户图片。
- 预检及插入前再次核对 DISPIMG 公式；不写公式、单元格数据或格式。
- 任务窗格启用三个图片操作按钮，保留公式设置固定开启，Convert 仍禁用。任务进行中防止重复点击。

## 验证

| 项目 | 结果 |
| --- | --- |
| TypeScript strict 类型检查 | 通过 |
| 自动测试 | 48 项通过，其中 22 项为本阶段新增 |
| Vite 生产构建 | 通过 |
| 微软 manifest 在线验证 | 通过 |
| 本机 HTTPS 任务窗格入口 | HTTP 200，正常证书校验 |
| 原生 Windows/macOS Excel 验收 | 待真实文件和宿主验证，未声称通过 |

新增测试包括横竖图比例、边界/隐藏尺寸、可选拉伸、名称与所有权、跨表和重复 ID、当前位置去重、合并区域、原生尺寸大图、保护工作表、插入前公式变化、刷新顺序、同步/异步宿主失败回滚和移除范围。

## 新增文件

- `src/core/image-layout.ts`
- `src/core/preview-identity.ts`
- `src/core/image-renderer.ts`
- `src/core/preview-controller.ts`
- `tests/image-layout.test.ts`
- `tests/image-renderer.test.ts`
- `tests/preview-controller.test.ts`
- `PHASE3.md`

## 修改文件

- `src/taskpane/taskpane.ts`
- `src/taskpane/taskpane.html`
- `package.json`
- `package-lock.json`
- `manifest.xml`
- `README.md`
- `public/support.html`

未增加任何 npm 依赖，未修改 WPS parser、工作簿读取或公式识别逻辑。

## 当前限制

1. Shape 预览需要 ExcelApi 1.10，旧版宿主保留扫描功能；合并区域完整适配需要 1.13。
2. 仅显示 PNG/JPEG。其他 MIME、损坏编码、超出宿主单次请求限制的图片报告失败或跳过，不自动转换格式。
3. 预览是会随文件保存的浮动 Shape，并非 Excel 原生单元格图片；原始公式一直保留。
4. 行列缩放和隐藏后的原生 TwoCell/比例锁效果需双平台实机验证。没有常驻事件监听器，尺寸改变后可用 Refresh 重新居中和适配。
5. Office.js 操作不是事务。生成失败会尽力回滚；若宿主断开或清理失败，可能留有临时图。删除旧图阶段失败可能暂时重复，修复权限后再 Refresh 或 Remove。
6. 扫描、资源导出和绘图不是原子快照。操作时请勿编辑/切换工作簿结构；代码会在写入前重新核对公式，但不能消除所有并发修改。
7. 删除仅遍历顶层 Shape。手工分组、移除/修改插件标记后不能保证自动识别；不会递归解组用户对象。
8. 真实 WPS 文件资源是否被 Excel 保留仍需验证；缺失资源不能恢复。Refresh 发现未解析资源会保留现有预览并停止。
9. 受保护工作表保守跳过，不修改保护设置。Refresh 涉及受保护工作表时停止并保留旧预览。
10. Convert 与 Cloudflare 部署未实现，本阶段不进入 Phase 4。

## 官方 API 依据

- [Excel Shape 创建与管理](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-shapes)
- [Excel Shape API](https://learn.microsoft.com/en-us/javascript/api/excel/excel.shape?view=excel-js-latest)
- 当前安装的 `@types/office-js` 标明 addImage / lockAspectRatio 为 ExcelApi 1.9，Shape.placement 为 1.10。
