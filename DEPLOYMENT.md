# Cloudflare Workers 部署

目标 Worker：`wpsimagecompat`
目标域名：https://wpsimagecompat.fogce.workers.dev
生产分支：`main`

## 本次错误原因

原项目只有 Vite 构建，没有 Wrangler 配置。Wrangler 在部署时检测到缺少配置，会进入自动配置流程并尝试修改 Vite；当前 async 配置不包含该流程所需的 plugins 数组，因此报 `Cannot modify Vite config: could not find a valid plugins array`。

修复采用 Workers Static Assets：仓库显式声明 `assets.directory = ./dist`，不需要 Worker 业务脚本或 Cloudflare Vite 插件。保留原 Vite 配置和 Office 本地 HTTPS 行为。

依据：[Cloudflare Workers Builds 配置文档](https://developers.cloudflare.com/workers/ci-cd/builds/configuration/)。

## Cloudflare 控制台设置

进入 `wpsimagecompat` → Settings → Build：

| 设置 | 值 |
| --- | --- |
| Git 仓库 | rich0022/WPSImageCompat |
| 生产分支 | main |
| 根目录 | 仓库根目录 |
| 构建命令 | npm run build |
| 部署命令 | npx wrangler deploy |
| Node.js | 22.12+，建议受支持的 LTS |

保存设置后重新部署**包含本次修复的最新提交**。不要重试锁定旧提交的失败构建；不要使用自动脚手架命令重新初始化 Vite。Workers 项目不需要填写 Pages 的输出目录字段。

构建流程：TypeScript → Vite → 验证并生成生产 manifest。部署流程：Wrangler 根据固定配置上传 dist 静态资源。compatibility_date 按 UTC 选择，避免亚洲本地日期已进入次日但 Cloudflare 仍认为日期在未来。

## Manifest 和路由

- 根目录 `manifest.xml`：localhost 开发版，不直接上传为生产文件。
- `scripts/prepare-deployment.mjs`：生成 `dist/manifest.xml`，将全部 localhost URL 替换为固定生产域名，并验证 XML 及资源存在。
- `/`：安装介绍页；通过 `_redirects` 重写到 `/index.html`。
- `/manifest.xml`：生产清单下载，application/xml，无缓存。
- `/src/taskpane/taskpane.html`：Excel 任务窗格，保留确切 .html 路径且无缓存。
- `/support.html`：帮助页。
- `/assets/icon-16.png`、`icon-32.png`、`icon-80.png`：Ribbon 图标。

禁用默认 HTML URL 改写，避免 manifest 中的 .html 入口被规范化到其他路径；未知路径返回 404，不把缺失资源伪装成任务窗格。

## 本地验证与发布

```sh
npm ci
npm test
npm run deploy:check
npm run validate:manifest:production
```

`deploy:check` 只构建并 dry run，不上传；生产清单验证使用微软在线服务。开发预览可运行 `npx wrangler dev --local --port 8787`，只能检查静态资源，不模拟 Excel 宿主。

如果由本机发布：

```sh
npx wrangler login
npm run deploy
```

使用 Git 自动构建时通常不需要本机登录，也不需要在源码中放置 Cloudflare token。

## 部署后验收

1. 首页、manifest、任务窗格、帮助页、三种图标都能通过 HTTPS 获取。
2. 下载的 manifest 中不能含 localhost，所有资源均指向正式域名。
3. 任务窗格返回 text/html，不应带阻止 Excel 嵌入的 X-Frame-Options / CSP frame-ancestors 限制。
4. 在 Excel 中换用生产 manifest，停止本地开发服务器后仍能打开插件。
5. 实际 Show/Refresh/Remove 行为仍按 README 的 Windows/macOS 真实文件矩阵验收。网页可访问不等于 Excel 兼容性通过。
6. `/version.json` 应返回 JSON 和 `Cache-Control: no-store`；其中 buildId 必须与本次打包的任务窗格脚本一致。HTML 使用 no-cache，带哈希的脚本文件随构建更新。不要单独发布版本文件，必须原子发布整个 dist。
7. 0.5.0 起，已有窗格在发现不同构建 ID 后提示用户刷新，忙碌时禁用更新。旧版本第一次升级须手动重开窗格。XML 清单权限或 Ribbon 更改仍需单独更新清单。

## 依赖说明

Wrangler 为开发部署依赖，不打包进 Excel 任务窗格。为修复其间接依赖已报告问题，将 Miniflare 的 sharp 限定到 0.35.4；这不用于插件图片处理。当前 adm-zip 0.6.0 在开发用 manifest 校验工具中仍有一项符号链接文件提取风险的审计告警（连带工具计两项 moderate），没有可用更新版；本项目只向该工具传入自有 XML 清单，不调用 ZIP 解包功能。运行时工作簿解包使用 JSZip。
