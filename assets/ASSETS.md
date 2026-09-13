# assets/ 通用前端库

本目录是 Access 数据库应用系统开发技能自带的**通用前端库**，用于在 Access 的「Web 浏览器控件」里渲染界面、内嵌仪表盘、做图表与交互页面。它们与任何自研框架无关，可独立复制进任意 Access 工程使用。

> 完整性校验：见 `assets/manifest.json`（逐项 sha256）。

## 清单与用途

| 文件 | 用途 |
|---|---|
| `css/bootstrap.min.css` | Bootstrap 3 栅格/组件样式，做响应式布局与表单/按钮/卡片 |
| `css/font-awesome.min.css` | Font Awesome 图标字体，配 `fonts/` 使用（图标以 `<i class="fa ...">` 引用） |
| `js/jquery-3.1.0.min.js` | jQuery，DOM 操作与事件的基础依赖 |
| `js/bootstrap.min.js` | Bootstrap 的 JS 组件（下拉、标签、模态框等），依赖 jQuery |
| `js/echarts.min.js` | ECharts 图表引擎，做折线/柱/饼/仪表盘等可视化 |
| `js/echarts-walden.js` | ECharts 的 walden 主题，统一图表配色 |
| `js/jquery.slimscroll.min.js` | 区域滚动条美化（可选） |
| `js/jquery.uniform.standalone.js` | 表单控件（输入框/下拉/复选）样式统一（可选） |

## 在 Access 里怎么用

1. 把这些文件随工程一起发布到某个本地目录（或打包进附件表、运行时写出到临时目录）。
2. Access 窗体放一个 `WebBrowser` 控件，用 `Navigate` 指向本地 `html`；HTML 里用相对路径引用上面的 `css/` 与 `js/`。
3. 数据通过 Access VBA 以 JSON 注入网页，或由网页回调 VBA（Access 的 `WebBrowser` 支持 `window.external` 双向通信）。

> 注意：这些是第三方通用库，本技能只引用、不修改其源码；如需升级版本，直接替换文件并更新 `manifest.json` 的 sha256 即可。
