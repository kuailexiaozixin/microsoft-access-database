# Edonsoft Development Framework_x64（占位）

本目录**未随本仓库分发**，仅作占位说明。

## 内容
Edonsoft 开发框架（64 位）本地副本，当前为上游 **v202610.01（2026-10-09 发布）**：`EdonSoft Development Framework.accdb`（主库）、`Data.accdb`、`edonsoftCore.accde`（编译核心）、`edonsoft-docs/`（官方帮助页 18 个栏目的完整 Markdown 转写：主题文档 `01-quickstart.md`~`18-date-picker.md`、总纲 `framework-overview.md`、架构 `13-architecture-source.md`、公开函数索引 `11-api.md` 与 `api/` 12 份分类正文，登记 80 个公开函数）、`appsettings.json`、`assets/`（含 v202610.01 新增的 `treeview`、`datatable` 两组 Web 控件资源）。

## 官方来源
- 官网：<https://www.edonsoft.com/>
- 框架产品页：<https://www.edonsoft.com/access-framework>
- 框架帮助文档：<https://www.edonsoft.com/access-framework-help.aspx>

## 在技能中的用途
作为 Access 快速开发的现成底座之一（本技能两大底座：Edonsoft 框架 + 盟威平台），提供表/窗体/菜单/权限等基础组件与官方文档，供业务系统开发复用。

## 换机使用注意（2026-10-15 真机实测）
- 包内 12 张链接表与 `appsettings.json` 的 `DbConnect` 指向上游作者机器的绝对路径，换到新机器直接打开会弹「后端数据库连接失败」；先用 DAO 重指到本地 `Data.accdb`（`td.Connect = ";DATABASE=<路径>"` 后 `RefreshLink()`，无需独占打开），重指后提示消失。
- 未完成登录初始化时，`Application.Run` 调用公开函数会抛 `disp_E_EXCEPTION`；先完成登录再复测。
- `USysFrmAccessTreeView` / `USysFrmAccessDataTable` 不在发布包内，是 `BuildAccessTreeViewForm` / `BuildAccessDataTableForm` 的生成产物，需现场生成。
- 对象规模（v202610.01 前端库枚举）：56 窗体 / 10 模块 / 1 宏 / 0 报表 / 24 表定义 / 105 查询定义，其中 102 个 `~sq_` 查询是上游作者另一套项目的宏临时查询，不属框架能力。
