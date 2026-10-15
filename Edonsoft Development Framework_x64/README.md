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
- **就地能打开能写，但可写打开会改写包内二进制**：本条先前记录的「技能目录这类路径上 Access 把前端库以只读会话打开（本地表 `Updatable=False`、开机弹 `错误 #3027 不能更新` 与 `#2471 … FAutoNum`、登录随之失败）」经 2026-10-15 原地复测**推翻**——就地打开时 `TS_TreeView`（本地表）与 `TS_Parameters`（链接表）`Updatable=True`、开机 0 弹窗、`admin/admin` 登录可用；卷与 8.3 短名策略、非 ASCII 目录名、路径长度、重解析点、DOS 只读属性、目录写权限、残留 `*.laccdb`、`Zone.Identifier`、他进程占用、双实例并存、DAO 独占后端、强杀脏状态逐轴实测排除，差异不在这个位置。换到工作目录运行的理由改成另一条实测：Access 以可写方式打开过一次，`EdonSoft Development Framework.accdb` 的字节就被就地改写（大小恒 5,509,120、sha256 逐轮变化），上游跟踪的内容基线会跟着漂移；DAO 只读打开（`DAO.DBEngine.120` 的 `OpenDatabase(path, False, True)`）不改写。所以核对包内事实走只读，真要跑改动就复制到普通工作目录，或跑完重新登记基线。
- 包内 12 张链接表与 `appsettings.json` 的 `DbConnect` 指向上游作者机器的绝对路径，换到新机器直接打开会弹「后端数据库连接失败」；用 DAO 重指到本地 `Data.accdb`（`td.Connect = ";DATABASE=<绝对路径>"` 后 `td.RefreshLink()`，无需独占打开）。技能目录内的这份已在 2026-10-15 完成重指，属本技能对 vendored 原件的有意本地化补丁。
- `Application.Run` 抛 `DISP_E_EXCEPTION`（`0x80020009`）与登录无关，是 pywin32 缺 makepy 类型信息时的封送失败：改用 `win32com.client.gencache.EnsureDispatch("Access.Application")` 早绑定句柄（`Run` 返回 `(返回值, 实参回显, …)` 元组，函数返回值取首元素），或纯 IDispatch 下走 `Application.Eval("函数名(参数)")`。开机未登录即可调用；登录后实测 `SQLText("O'Brien")` → `'O''Brien'`、`HashPassword("abc")` → `v2$1200$…`、`ConvertToJson("abc")` → `"abc"`、`AutoNumStr("tblChatHistory","ID",4,"X","yyyymmdd")` → `X202610150001`、`GetUnreadCount()` → 1。默认账号 `admin/admin`（`VerifyPassword("admin", <TS_Users.FPassword 原值>)` → True）。
- 登录只有界面路径（官方文档未给可编程登录函数，`api/07-security.md` 仅 `LogOff`）。自动化下的触发方式：`Run "LoginSub_OKClick", Forms("USysFrmLogin").Controls("sfrChild").Form`，形参是该子窗体对象；`SetForegroundWindow` + `keybd_event` 与 `DoCmd.SendKeys` 在此均不可靠。
- `USysFrmAccessTreeView` / `USysFrmAccessDataTable` 不在发布包内，是 `BuildAccessTreeViewForm` / `BuildAccessDataTableForm` 的生成产物，需现场生成。
- 对象规模（v202610.01 前端库枚举）：56 窗体 / 10 模块 / 1 宏 / 0 报表 / 24 表定义（12 本地 + 12 链接）/ 105 查询定义，其中 102 个 `~sq_` 查询是上游作者另一套项目的宏临时查询，不属框架能力。
