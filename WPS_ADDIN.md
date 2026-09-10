# WPS 表格加载项

`wps-image-compat-et` 是 WPS Image Compat 面向 **WPS 表格（ET）** 的独立 Web 加载项。它不使用 Excel 的 XML manifest，也不会读取或上传工作簿文件。

## 安装

1. 在安装了桌面版 WPS 表格的 Windows 或 Linux 电脑上，用 Chrome、Edge 或系统默认浏览器打开 [WPS 加载项安装页](https://wpsimagecompat.fogce.workers.dev/wps-et/publish.html)。
2. 在 `wps-image-compat-et` 这一行点击启用；页面显示“配置成功”后关闭页面。
3. 完全退出并重新打开 WPS 表格，在 **WPS Image Compat → 图片与兼容工具箱** 打开任务窗格。

安装页由 WPS 的发布工具生成，通过本机 WPS 服务将线上地址登记到当前用户的 WPS 中。每台电脑需要各自启用一次；后续任务窗格代码更新由线上地址加载，无需重新安装。若 manifest、功能区或插件标识变更，则应重新打开安装页执行更新。

## 当前功能

- 只读扫描全部工作表中的 `DISPIMG` / `_xlfn.DISPIMG` 公式，列出位置和图片 ID。
- 读取 WPS 可公开提供 `ImageUrl` 的普通图片 Shape，显示其工作表和左上角单元格位置。
- 由用户点击“选中首张图片”时，显式定位到本次扫描到的第一张普通图片。

扫描不写入公式、单元格、Shape 或工作簿设置。它用于确认 WPS 内的图片和公式分布，供后续的 Excel 兼容处理使用。

## 已知范围

- WPS 官方加载项文档当前面向 Windows 和 Linux；本项目没有声明 macOS WPS 支持。
- “图片置入单元格”在 WPS JS API 中没有可稳定读取的跨版本接口。当前只统计暴露 `ImageUrl` 的普通图片 Shape，不会把未公开给 API 的图片误报为丢失。
- 此版本不转换 DISPIMG、不删除图片，也不改写 Excel 或 WPS 公式。
- 无法看到功能区时，先确认安装页的条目状态正常，完全重启 WPS；企业/OEM 环境可能由管理员限制加载项渠道。

## 发布者检查

线上文件应可直接访问：

- `https://wpsimagecompat.fogce.workers.dev/wps-et/ribbon.xml`
- `https://wpsimagecompat.fogce.workers.dev/wps-et/index.html`
- `https://wpsimagecompat.fogce.workers.dev/wps-et/ui/taskpane.html`

不要通过修改个人 WPS 的 `oem.ini` 手工加载本插件；本项目使用 WPS 的在线发布/安装流程。组织如需统一分发，应按 WPS 的企业或 OEM 加载项渠道配置。
