# Phase 1 交付记录

日期：2026-09-08。仅实现 Phase 1，未进入 Phase 2–4。

## 验证结果

| 检查 | 结果 |
| --- | --- |
| TypeScript strict 类型检查 | 通过 |
| 自动测试 | 5 项通过：公式识别、排除无效表达式、ID 转义、A1 坐标、模拟 Excel 扫描 |
| Vite 生产构建 | 通过；输出 dist/ |
| 微软 manifest 在线验证 | 通过，输出 The manifest is valid |
| npm 安装安全审计 | 0 vulnerabilities |
| 本地开发证书 | 已生成并安装信任 |
| HTTPS 任务窗格入口 | HTTP 200，未跳过证书校验 |
| macOS 侧载准备 | 已复制 manifest 至 Excel wef/wps-image-compat.xml，并核对一致；已启动 Excel |
| macOS 原生 Ribbon / 工作簿扫描 | 待人工验收；当前工具不能检查原生 Excel 界面 |
| Windows 原生加载 / 工作簿扫描 | 待 Windows 环境验收 |

模拟测试不是实际宿主验证，manifest 验证只证明清单有效。开发服务器运行于 https://localhost:3000；会话结束后若服务停止，重新执行 npm run dev。

## 新增文件清单

- `.gitignore`
- `package.json`
- `package-lock.json`
- `tsconfig.json`
- `vite.config.ts`
- `manifest.xml`
- `src/types/wps.ts`
- `src/core/dispimg-formula.ts`
- `src/core/dispimg-scanner.ts`
- `src/taskpane/taskpane.html`
- `src/taskpane/taskpane.ts`
- `src/taskpane/taskpane.css`
- `public/assets/icon-16.png`
- `public/assets/icon-32.png`
- `public/assets/icon-80.png`
- `public/support.html`
- `tests/dispimg-formula.test.ts`
- `tests/dispimg-scanner.test.ts`
- `README.md`
- `PHASE1.md`

生成目录 node_modules/ 与 dist/ 已忽略。开发证书与 Excel 侧载文件位于本机用户目录，不属于仓库。

## 当前限制与后续问题

- 双平台原生 Excel 验收尚未完成；请按 README 的测试矩阵执行。
- Phase 1 只识别直接 DISPIMG 调用、字面量 ID、模式 1。嵌套公式和引用参数暂不支持。
- 分批扫描有值使用区域，极稀疏但跨度很大的工作表仍可能较慢；扫描期间应避免修改工作簿，避免分批结果不一致。
- 资源解析未实现，parsed/missing 显示尚未检查。没有读取或修改 OOXML。
- Shape、刷新、删除和转换按钮禁用，设置保持预选但暂不能修改。
- 官方验证工具的间接 adm-zip 依赖经 scoped override 升至 ^0.6.0，以解决安装发现的漏洞；在线清单验证已在覆盖后通过。
- manifest 为满足微软校验使用版本 1.0.0.0；npm 项目仍为 Phase 1 版本 0.1.0。
- GitHub 远程仓库未创建、未推送、未发布。当前仅使用用户指定的仓库名称。
