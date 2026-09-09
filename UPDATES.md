# 0.5.0：中英界面与发布更新提示

任务窗格提供简体中文/英文和跟随 Excel 选项，翻译按钮、设置、状态、统计标签和常见错误；原始宿主错误可展开查看。翻译字典与 Office 操作解耦，未新增运行时依赖。语言选择只保存在本机当前站点，工作簿内容不翻译、不上传。

每次生产构建创建独立构建 ID，写入任务窗格脚本与 version.json。运行中的生产窗格打开、获得焦点及可见时每 5 分钟检查版本，无凭据、无工作簿请求数据。发现变化后提示用户刷新，操作期间禁用按钮；页面刷新后需要重新扫描。开发环境继续使用 Vite；生产不使用代码注入或 Service Worker 强制热替换。

## 修改文件

- `src/taskpane/i18n.ts`：纯函数语言选择、字典、参数替换与错误码翻译。
- `src/taskpane/taskpane.ts`、`taskpane.html`、`taskpane.css`：语言选择、状态重译、更新提示与忙碌状态保护。
- `src/core/update-checker.ts`：同源版本检查、超时、离线容错与重叠检查合并。
- `src/core/preview-controller.ts`、`image-renderer.ts`：向 UI 提供阶段和错误码，保留原始错误详情。
- `src/types/build.d.ts`、`vite.config.ts`、`scripts/prepare-deployment.mjs`、`public/_headers`：构建标识、版本文件和缓存策略。
- `package.json`、`package-lock.json`、`tsconfig.json`：版本号及 DOM 集合类型支持，没有增加依赖。
- `tests/i18n.test.ts`、`tests/update-checker.test.ts`：语言回退、翻译参数完整性、版本变化、离线、超时和更新保护测试。
- `README.md`、`TESTING.md`、`DEPLOYMENT.md`、`public/index.html`、`UPDATES.md`：使用、发布与验收说明；Mac 入口按用户实机确认改为“开始 → 加载项”。

TypeScript、68 项自动测试和构建通过。本地浏览器已验证中英切换、重开后的语言记忆、模拟构建变化时的更新提示以及点击刷新；测试用版本信息已恢复。Vite preview 不再触发开发证书安装。此前 Phase 4 变更仍见 PHASE4.md。原生 Excel 中的语言切换、刷新后的图片保留，以及 Windows/macOS 的完整图片流程仍需验收。

## 范围

首版支持简体中文和英文；其他中文地区回退简体，其他语言回退英文。产品名称和 Ribbon 保持原名。更新提示只在生产构建启用；第一次从旧版升级需重开窗格；XML 清单本身不由这套机制替换。刷新窗格不等于转换工作簿。
