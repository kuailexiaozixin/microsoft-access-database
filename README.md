# Microsoft-Access-DataBase（Access 数据库应用系统开发技能）

本技能帮你在 Microsoft Access 里把一套业务系统做出来、做对、并能长期维护。核心主张是：**能复用的成熟积木就复用，不要重复造轮子；版本控制（VCS 加载项 + msaccess-vcs-mcp）是贯穿开发全生命周期的方法，不是收尾才做的单独一步；每个阶段都有明确的输入、动作与输出物**。本目录自带经过验证的积木：两大现成底座（各带完整官方文档）、通用控件与代码库、AI 能力、版本控制三通道、以及 `examples/` 下的完整可运行案例。

> 本技能最终会作为 `../VBA/SKILL.md` 的子技能（`access-vba-addin-builder/`）迁移进去。当前它以独立目录形态存在，所有能力可直接使用。

## 这条技能能干什么

- 用 Access + VBA 从零或基于现成底座构建业务系统（表/关系/查询、窗体/菜单/主界面/报表、VBA 逻辑）。
- 接入大模型（在 Access 内直接调用 AI 问答、Markdown 渲染、基于文本源文件做 AI 协作开发）。
- 用 Version Control System 加载项 / MCP 把 `.accdb` 导出成文本源文件、做源码版本管理与构建回库——这是贯穿每个阶段的方法。
- 用 Python + win32com 做零手动的自动化测试、验证、漂移核验与排错。

## 9 阶段工作流（详见 `SKILL.md`）

环境准备与底座选型 → 需求分析与系统设计 → 数据建模与查询 → 通用代码库（模块/类模块/宏）→ 窗体与界面 → 报表设计 → AI 能力接入 → 自动化测试与验证 → 部署与运维。版本控制在「环境准备」阶段就位，从「需求分析」起把库导出成文本源 `.src` 作为单一真相源，之后每一步的变更都走「改 `.src` 文本 → 构建/回写回库 → 漂移核验」的闭环，每个阶段写明任务内容、任务要求与输出物。每一步该用哪个积木、怎么用、VCS 何时导出何时构建，写在 `SKILL.md` 里，是一条连续的路而不是菜单。

## 自带积木（底座、案例与工具）

**现成底座（含完整文档）**
- `盟威Access快速开发平台V2.7.0版(64位)/`：企业级快速开发平台，含登录、权限、动态菜单、主界面等开箱即用能力（`.accdb`/`.accde` + 后端 `Data.accdb`）；完整开发教程见其自带文档 `盟威Access快速开发平台V2.7.0版(64位)/Access2016报销管理系统案例开发教程.pdf`（平台概述 → 程序设计文档 → 表/查询/窗体/报表 → 用户权限设计，九章完整示范一个报销系统的开发全过程）。
- `Edonsoft Development Framework_x64/`：同类企业框架（前端宿主库 + 核心编译库 + 后端 `Data.accdb`），18 项内建能力；官方帮助页全部栏目已完整转写为自带文档 `Edonsoft Development Framework_x64/edonsoft-docs/`（总纲 `framework-overview.md`、架构 `13-architecture-source.md`、功能主题 `01-quickstart.md` ~ `12-faq.md`、60 个公开函数见 `Edonsoft Development Framework_x64/edonsoft-docs/api/`）。

**通用 UI 控件（`examples/` 下，可复用组成部分）**
- `examples/Access DatePicker/README.md`：纯 VBA、零依赖的日期/年月选择器（`CreateDatePickerForm` / `ShowDatePicker` / `ShowYearMonthPicker`），含演示库与 `使用说明.md`。
- `examples/Access VBA Modules Collection/README.md`：高频能力模块集合（自动编号 `basAutoNumStr`、字段校验 `ClsFieldValidator`、ADO 数据访问 `ADOExecute`、导出 Excel/PPT/HTML、VBE 工具 `modVBETools`），含 `examples/Access VBA Modules Collection/wiki/` 文档。

**AI 能力**
- `examples/AccessAI/README.md`：在 Access 内调用大模型的成熟实现（`JsonConverter.bas` / `Module_Markdown.bas`，`CreateAIForm` / `ConfigureApiKeys` / `ShowMarkdown`）。
- AI 文本化开发范式（导出文本 → 编辑 → 构建回库）已内化为本技能工作流（见 `SKILL.md` 总览「统一闭环」与阶段 6.2）；其编辑规则与对象级手册为本技能 `references/` 自有文件——通用规则（对象配对、编码换行、安全删除）见 `references/vcs-text-editing.md`，9 份对象级手册（`forms.md`/`reports.md`/`images.md`/`conditional-formatting.md`/`queries.md`/`tables-and-relationships.md`/`vba.md`/`project-config.md`/`other-objects.md`）按该文件第六节映射表取用；完整样例与 prompt 模板见 `examples/CustomerOrders/`。

**业务样例（`examples/` 下）**
- `examples/Access BOM Management System/README.md`：多级 BOM 管理样例（树形视图 + 父子关系维护），层级结构数据的落地参考。
- `examples/01_blank_db_to_vcs_loop/` 与 `examples/02_edit_src_reimport/`：两个零手动闭环演示（空库 → 导出 `.src` → 漂移核验；改 `.src` → 回写 → 复核）。

**版本控制通道（四实体 + 技能目录自包含运行位，底层都调同一个加载项引擎）**
- `msaccess-vcs-addin/README.md`：加载项**源码仓库**（源码 + 文档 + `Ribbon\Build\` 官方编译的 Ribbon COM DLL）。Version Control.accda 既可在 Access 里用 Ribbon「Export / Build / Merge」，也可被 `msaccess-vcs-mcp` 与 `scripts/vcs_drive.py` 以代码方式驱动——本技能所有出入库动作都走代码/MCP 路径，不依赖人工点按。
- `Microsoft Access Version Control System/`：加载项**安装位=运行位（技能目录内，自包含）**——`Version Control.accda`（修复版本体）+ `MSAccessVCSLib_win64.dll`（Ribbon COM 加载项）+ `Ribbon.xml/json` + `Worker.vbs`；Menu Add-Ins、受信任位置、CLSID 全部指向这里。**操纵一律优先此运行位**（副本损坏是加载项崩溃的首要根因）。历史旧安装位 `%AppData%\MSAccessVCS\` 已回收清理（移入回收站可还原），不再参与运行。
- `Version_Control_v5.0.1/`：**v5.0.1 官方安装程序**（含内嵌资源，代码化安装的源输入；留档，常规开发不用）。
- `msaccess-vcs-mcp/README.md`：Python 写的 MCP 服务，把加载项包装成 `vcs_export_database` / `vcs_import_object` / `vcs_rebuild_database` / `vcs_compile_vba` 等工具，供 AI 编码代理自动驱动（自包含安装配置见 `references/vcs-three-channels.md` 1.2 节）。

## 配套文档与工具

- `SKILL.md`：唯一的内容入口，以工作流形态呈现，链接嵌入论述（无索引表/路由表/资源清单）。
- `docs/troubleshooting.md`：按症状查的排错手册（故障 → 根因 → 修复）。
- `references/`：硬规则与方法论（编码铁律、工程操作、VCS 通道与安装配置、`.src` 文本编辑规则、架构取舍、自动化 SOP、完整案例 PDF）；分工与导航见 `references/README.md`。
- `references/开发财务管理系统.pdf`：完整真实案例（分析 → 建表 → 查询 → 窗体 → 报表 → 代码模块 → 启动），章节印证工作流各阶段。
- `references/谈数据库设计前分析工作【Access软件网】.html`：数据库设计前的分析工作（只读参考）。
- `examples/`：实操案例（成型系统或者可复用的组成部分）——照着真机跑一遍就能看到「改 `.src` → 构建/回写 → 漂移核验」闭环的具体场景（每个案例附可运行脚本）。
- `templates/`：可复用的模块/类模块/宏源码模板（已去框架化）。
- `scripts/`：零手动的 Python 可运行脚本与代码示例（弹窗守卫、回收站删除、漂移核验、模块重建、VCS 探针与驱动、COM 连接等可复用代码）。
- `tests/`：通用验证脚本（模板/编码/漂移核验的冒烟测试）。
- `assets/`：通用前端库（Bootstrap、jQuery、ECharts、Font Awesome），用于 Access 内嵌网页/仪表盘。

## 怎么开始

1. 读 `SKILL.md` 走一遍 9 阶段，确定你的底座（盟威 / Edonsoft 二选一）与要复用的积木。
2. 动手前先读 `references/iron-laws.md` 把编码、类模块、工程定位、弹窗守卫这几条刻进习惯。
3. 想照着真机跑一遍闭环，看 `examples/` 里的具体示例。
4. 遇到问题按症状去 `docs/troubleshooting.md` 查；三套 VCS 工具的正确姿势读 `references/vcs-three-channels.md`。


## 第三方资源（占位说明）

以下目录包含第三方框架/安装包/源码仓库的本地副本，**未随本仓库分发**，以 `README.md` 占位并链接官方来源：

| 目录 | 官方来源 |
|---|---|
| `Edonsoft Development Framework_x64/` | <http://www.edonsoft.com/access-framework-help.aspx> |
| `Microsoft Access Version Control System/` | <https://github.com/joyfullservice/msaccess-vcs-addin> |
| `Version_Control_v5.0.1/` | <https://github.com/joyfullservice/msaccess-vcs-addin/releases> |
| `msaccess-vcs-addin/` | <https://github.com/joyfullservice/msaccess-vcs-addin> |
| `msaccess-vcs-mcp/` | <https://github.com/joyfullservice/msaccess-vcs-mcp> |
| `盟威Access快速开发平台V2.7.0版(64位)/` | <http://www.accessgood.com/> |

克隆/下载方式见各占位 `README.md`。
