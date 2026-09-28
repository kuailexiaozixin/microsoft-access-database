---
name: microsoft-access-database
description: Microsoft Access 数据库应用系统开发技能（通用、不绑定任何自研框架）。用 Access + VBA 从零或借鉴现成框架与积木（盟威平台 / Edonsoft 框架）构建业务系统：需求分析与系统设计先行、数据建模与查询、通用代码库（模块/类模块/宏）、窗体与控件、报表、AI 能力接入、自动化测试与验证、部署与运维。内容按开发生命周期组织为 0–8 九个阶段，每阶段分工作步骤，每步四要素（任务内容 / 任务要求 / 输出物 / VCS 融入）。版本控制（VCS 加载项 + msaccess-vcs-mcp）作为贯穿全程的方法，在每个阶段的具体环节落地（何时导出、何时构建、何时核验、用什么脚本、输出什么）。当需要设计 Access 表/关系/查询、编写或注入 VBA、构建窗体/菜单/主界面/报表、接入大模型、用 VCS 对 .accdb 做源码导出/构建/回写、排查「找不到过程」「中文乱码→SQL 3075」「类模块 type 退化」「加载项崩溃」「弹窗卡死」等 Access VBA 问题时使用本技能。复用本目录的成熟积木（框架积木、通用控件、通用代码库、AI 能力、版本控制通道、完整案例），而非从零造轮子。
---

# Access 数据库应用系统开发工作流

## 一、总览：怎么读本文档

**一条主线**：输入是业务想法，输出是三样东西——可部署的前端副本、可审计的文本源（`.src`）、能自动验证的行为。中间由 0–8 九个阶段连续执行，前一个阶段的输出物就是后一个阶段的输入，每步统一回答四件事：

- **任务内容**——做什么；**任务要求**——按什么标准做（含格式手册、样例、脚本）；**输出物**——做完留下什么；**VCS 融入**——版本控制在本步怎么用（仅触及库对象文本的步骤有差异动作；设计文档期与纯数据期无差异动作则不出现该字段）。

**六段推进**：分析设计（阶段 1）→ 数据与逻辑（阶段 2–3）→ 界面与输出（阶段 4–5）→ 增强（阶段 6）→ 验证（阶段 7）→ 交付（阶段 8）。阶段 0 是开工前的环境、框架与版本控制通道准备，可并行完成。

**三条方法主张**贯穿全程：

1. **能复用的成熟积木就复用，不重复造轮子**——框架积木、控件、代码库、AI 能力、完整案例都自带（见各阶段引用与 `examples/README.md`）。
2. **版本控制是贯穿开发全生命周期的方法，不是收尾才做的单独一步**——库一建好就导出 `.src` 作单一真相源（阶段 2.1 锁定基线），此后一切变更走统一闭环：**VCS 导出 `.src` → 在文本源上改 → 构建/回写回库 → 漂移核验**。各阶段 VCS 融入只写本步的差异动作。
3. **每个阶段都有明确的输入、动作与输出物**——不建无来源的库对象，不产出无去向的文档。

**架构立场**：本技能以**分层 + 面向对象建模 + 前后端物理分离**组织代码，各原则在阶段落地处见下：

- **分层（六段即分层）**：数据层（表/关系，阶段 2）→ 取数层（查询，阶段 2.4）→ 逻辑层（标准模块=业务函数、类模块=有状态对象，阶段 3）→ 交互层（窗体，阶段 4）→ 输出层（报表，阶段 5）→ 增强/验证/交付（阶段 6–8）；每层经下层接口取数，不越层直连（如窗体不直接写 SQL，取数走查询或逻辑层函数）。
- **面向对象建模**：业务实体/服务收进类模块（3.2），用属性/方法/事件表达行为——实体逻辑不散落在窗体事件里。
- **前后端分离（两层含义）**：物理层=前端程序 accdb + 后端数据（`Data.accdb` 链接表 / 升迁 SQL Server/MySQL，8.1）；界面层=WebBrowser 内嵌网页/仪表盘的前端资源（assets css/js，4.2/4.3）。
- **模块化 / 高内聚低耦合 / 关注点分离**：模块划分先行（1.3）→ 每层内按职责分模块（3.1 十目录分层范本）→ 布局与代码分离（4.1 `SplitLayoutFromVBA`：`.bas` 布局 / `.cls` 代码）→ 可测逻辑写成不依赖界面的纯函数（7.1）→ `.src` 文本源按对象类型分目录（对象全景表）。

**统一闭环的工具链**（阶段 2.1 详述，后续各步直接引用）：
- 导出用 `scripts/vcs_drive.py` 的 `FullExport` 子命令；构建/回写用同一脚本的 `Build` 子命令或 VCS MCP（`msaccess-vcs-mcp`）；
- 漂移核验用 `scripts/vcs_consistency_check.py`；
- 所有脚本统一套弹窗守卫 `scripts/fwguard.py`（捕获消息框文本 + 自动点击按钮 + 进程清理，杜绝手动弹窗）；
- 违反文本格式规则会静默损坏库或在回写时报 `Error 2128`（`LoadComponentFromText` 导入失败）。

**MCP 侧的完整操作手册**在 `msaccess-vcs-mcp\docs\`（上游仓库自带文档，不在本技能正文重复，各阶段按需取用）：
- `AGENT_WORKFLOWS.md`（AI 代理 8 种工作流）、`EXPORT_FORMATS.md`（导出文件格式）、`TESTING.md`（MCP 三层测试）
- `VBA_CALLBACK_API.md`（异步/取消/并发回调契约）、`VBA_INTEGRATION.md`（MCP↔加载项 COM 集成架构）
- **MCP 源码 `msaccess-vcs-mcp\src\` 是"用代码自动操作 Access"的权威实践样板**——写自己的自动化脚本前照它抄骨架：连接与实例生命周期、加载项调用链、21 个 `vcs_*` 工具参数与返回处理、进程登记与清理、COM 串行门禁、隔离 VBA 执行的落点见 0.3 与 7.2 手段 F（不再在此重复展开）。

**文档分工**：
- 动手前先读 `references/iron-laws.md` 把编码、类模块、工程定位、弹窗守卫刻进习惯；
- 出错按症状查 `docs/troubleshooting.md`；技能自身资源地图见 `README.md`；
- 变化只记在 `changelog.md`；
- 改 `.src` 文本的通用硬规则（配对、编码换行、安全删除、常见坑）见 `references/vcs-text-editing.md`（含按对象查手册的映射表）。
  - **VBA 语言与 Access 对象模型查证**：微软官方 VBA-Docs 资料库（`..\VBA-Docs\`，3.2 万+ API 页 + 概念手册）——写 VBA 代码/查对象模型/查 Access SQL 语法前先查它，各阶段按需直接引用其子目录（2.4 查询语法、3.1 DAO/ADO 与语言参考、3.2 类模块与错误处理、4.1/4.2 窗体控件、7.2 错误消息、8.3 启动属性），索引检索见 VBA 技能的 `..\vba_docs_index.db`。
  - **Access 对象模型 API 单页（`..\VBA-Docs\api\`，3702 个 `Access.*.md`，对象/方法/属性/事件逐成员一页，按名即查）**：
    - 总索引：`api\Access.md`（枚举总索引 AcXxx 全部枚举页）、`api\overview\Access.md`（Access VBA reference 应用概览）；
    - 按对象类取用：`Access.Application.*`（123 页：CurrentDb/CurrentProject/AccessError 等）、`Access.DoCmd.*`（66 页：OpenForm/OpenReport/RunSQL/OutputTo 等）、`Access.Form.*`（252 页）、`Access.Report.*`（178 页）、控件类（ComboBox 178/TextBox 163/ListBox 145/CommandButton 143/ToggleButton 148 等）、`Access.Module.*`（AddFromString/AddFromFile 等）；
    - DAO/ADO 对象模型本地无 api 页，用 `references/vba-database-programming.md` 官方原文汇编；各阶段按对象类型点名下节（3.1 应用/DoCmd/Module、4.1 Form、4.2 控件、5.1 Report、7.2 AccessError）。

**加载项官方 Wiki**（`msaccess-vcs-addin\Wiki\`，上游完整转储）是加载项侧的用户文档库——各阶段已按需直接引用（0.3 安装/配置/MCP 自动化、对象全景对象与文件类型、2.4 查询源文件、4.1 拆分文件、6.2 保存即导出、7.2/7.4 测试与合并构建、8.1 连接、8.3 发布、排错 FAQ），此处不再重复导航。

## 二、对象全景与文本形态

一个 Access 应用由七类核心对象加若干配套对象组成。**下表是"对象 → 处理阶段 → 文本形态 → 手册"的路由视图（供按对象查找手册，本身不是工作流步骤）**；工作流执行仍按阶段顺序推进——每类对象在本工作流中有固定的处理阶段，**改哪种对象、处理哪个环节，按手册读；下文各阶段直接点名手册与活样例，不重复正文**：

| 对象 | 处理阶段 | `.src` 文本形态 | 格式手册 |
|---|---|---|---|
| 表（本地/链接/系统） | 阶段 2 | `tbldefs/`：本地表 `.xml`（XSD），链接表 `.json` | `references/tables-and-relationships.md` |
| 关系（外键/引用完整性） | 阶段 2 | `relations/*.json` | `references/tables-and-relationships.md` |
| 查询（选择/操作/联合/传递/子查询） | 阶段 2 | `queries/`：`.bas`（元数据）+ `.sql`（SQL 文本），恒成对 | `references/queries.md` |
| 模块（标准 + 类） | 阶段 3 | `modules/`：`.bas`（标准模块）、`.cls`（类/代码后置） | `references/vba.md` |
| 宏（AutoExec 等启动动作） | 阶段 3 | `macros/*.bas`（Access 宏动作，非 VBA） | `references/other-objects.md` |
| 窗体（人机交互） | 阶段 4 | `forms/`：`.bas`（布局）+ `.cls`（代码后置）+ `.json`（打印设置） | `references/forms.md` / `images.md` / `conditional-formatting.md` |
| 报表（输出打印） | 阶段 5 | `reports/`：文件配对同窗体 | `references/reports.md` / `images.md` / `conditional-formatting.md` |
| 配套对象（导入导出规格、共享图像、主题、库属性、VCS 工程文件） | 随项目 | `imexspecs/`、`images/`、`themes/`、顶层 `.json` | `references/project-config.md` / `other-objects.md` |

- **硬规则总入口**：见 `references/vcs-text-editing.md`（11 种对象配对、UTF-8-BOM/CRLF 编码、新建文件显式写 BOM 后 `od -c -N 5` 验证、安全删除清单）；编码铁律导入侧细节（UTF-16 LE 窗体）见 `references/vba-encoding-rules.md`。
- **活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/` 是九份手册的配套完整样例（一个 prompt 生成，3 表 4 查询 6 窗体 1 报表），每种对象先读手册规则、再到示例对应文件核对格式。
- **更大规模生产级 `.src` 工程范本**：`msaccess-vcs-addin\Version Control.accda.src/`——VCS 加载项自身的完整源码工程（10 个模块分层 + 类模块封装 + 自带测试套件），阶段 3 写模块/类、阶段 7 写测试前可对照其工程组织。
- **Wiki 权威**：对象清单/导出文件类型/命名术语见 `msaccess-vcs-addin\Wiki\` 三篇（Supported-Objects / Export-Import-File-Types / Terminology-and-Style-Guide），写对象前可与 9 份手册互查。

## 阶段 0：环境准备（环境 → 框架与积木 → 版本控制）

开工前把三件事备齐：运行环境、框架与积木地图、版本控制通道。三者互不依赖，可并行完成。

### 0.1 环境就绪

- **任务内容**：确认 Microsoft Access 已安装且位数匹配（32/64 位与 Office 一致）；开启「信任对 VBA 工程对象模型的访问」；把工作目录加进「受信任位置」（否则 win32com 驱动与 AutoExec 失效）。
- **任务要求**：AutoExec 宏的源码模板见 `templates/autoexec-macro.macro`；环境是否就绪用空库探针 `scripts/probe_blank_db.py` 快速验证。
- **输出物**：环境检查结果（Access 版本、位数、信任设置、探针报告）。

### 0.2 框架与积木地图（两框架组成拆解）

- **任务内容**：把两大框架的组成拆解为可复用的积木清单，构建业务系统时按阶段按需取用。**不做"二选一"选型**——两框架的能力、教程、文档在后续各阶段直接引用、互不排斥，LLM 借鉴其组成与实践来构建自己的应用系统。

- **任务要求**：
  - **盟威 Access 快速开发平台**（`盟威Access快速开发平台V2.7.0版(64位)/`）：
    - **组成**：`Main.accdb`（客户端主程序）、`Data.accdb`（后台数据库）、`Update.accde`/`Download.accde`（自动升级）、`RDPLib.ucl`（支持库）、`Config.ini`（配置）。
    - **完整教程**（`Access2016报销管理系统案例开发教程.pdf`）九章对应本工作流阶段：
      - 第 3 章程序设计文档（需求/功能/表设计说明书）→ 阶段 1；第 4 章表设计 → 阶段 2.2
      - 第 5 章查询设计 → 阶段 2.4；第 6 章窗体设计（自动编号、数据模块向导、导航菜单、查阅列表、启动菜单）→ 阶段 4
      - 第 7 章报表设计（静态/动态/切换面板）→ 阶段 5；第 8 章用户权限设计 → 阶段 8.2
      - 第 2.4 节文件构成与自动升级 → 阶段 8.3
    - **源码与组成**：盟威库以加密/向导形态发布（`反编译打开Main.accdb.vbs` 提供 `/decompile` 入口），借鉴其组件划分与向导产物（自动编号、数据模块向导、导航菜单、查阅列表、启动菜单、报表向导、用户权限），不直接改库。
    - **可见可读部分**：`Config.ini`（登录/主题/自动登录配置格式，8.2 引用）、`Images/`（登录界面与主题图片：login_classic/login_standard/head.jpg/progress.jpg，4.2 界面资源参考）、教程 PDF（九章完整案例）。
  - **Edonsoft 开发框架**（`Edonsoft Development Framework_x64/`）：
    - **组成**：`EdonSoft Development Framework.accdb`（前端宿主库）、`edonsoftCore.accde`（核心编译库）、`Data.accdb`（数据后端）、`appsettings.json`（连接配置）、`assets/`（主题/前端资源）。
    - **能力与源码落点已按阶段融入工作流各步骤**（各步骤处均给出真实窗体/模块/函数落点，已枚举核实）：
      - 登录/RBAC → 8.2；通用查询 → 2.4；公共函数/操作日志 → 3.1；代码生成器 → 4.1
      - 搜索组合框/单据选择/进度条 → 4.2；动态菜单/待办/仪表盘 → 4.3；AccessAI → 6.1
      - 链接表/升迁/多后端 → 8.1；自动更新/设置 → 8.3
      - 完整落点清单见 `13-architecture-source.md` 第 8 节表格；框架库本体以 accde/加密发布，借鉴其落点与划分，不直接修改。
    - **文档**：设计理念见 `framework-overview.md`；架构详述（前后端分离、三后端、RBAC、动态菜单 `FArguments`、`mode/` 源码目录、32/64 位）见 `13-architecture-source.md`；其余能力按官方帮助页左侧菜单完整转写为 `01-quickstart.md`~`12-faq.md` 与 `api/`（60 个公开函数，7 类）。
    - **可见可读部分**：`framework-overview.md` 第 4 节为 COM 枚举核实的对象清单（AllForms/AllModules）与可见源码模块清单（用途与可移植见 3.1）；`appsettings.json`（多后端连接配置 ActiveBackend/Backends[]，8.1 引用）；`assets/`（css/js/images/Theme 前端资源，4.2 参考）。
  - **两框架差异对照**（各能力的实现方式、文档形态、借鉴场景对比）见 `references/framework-comparison.md`——它是**对照参考不是选型决策**：两框架能力有重叠时，按对照表选实现更成熟或更贴近业务形态的借鉴。
- **输出物**：框架积木地图（能力 → 阶段 → 文档 → 用法，见上两段映射）。

### 0.3 版本控制通道就位

- **任务内容**：安装并配置版本控制通道——加载项本体 + MCP 服务。
  - **安装位与运行位统一在技能目录内 `Microsoft Access Version Control System\`**（自包含、自运行，不注册/安装到技能目录之外）；
  - 四个目录实体的关系（源码仓库 `msaccess-vcs-addin\` / 完整分发部署包 `Version_Control_v5.0.1\` / 运行位 `Microsoft Access Version Control System\` / MCP 服务 `msaccess-vcs-mcp\`）、加载项与 MCP 的代码化安装配置、三种操作通道（Ribbon / API / MCP）怎么选，全在 `references/vcs-three-channels.md` 第一章与第二章。
- **任务要求**：
  - **加载项（全代码化安装，禁止手动弹窗操作）**：运行 `python scripts/install_vcs_addin.py --source <构建好的 accda> --folder <技能目录\Microsoft Access Version Control System> --ribbon-dll <msaccess-vcs-addin\Ribbon\Build\MSAccessVCSLib_win64.dll> --ribbon-xml <msaccess-vcs-addin\Ribbon\Ribbon.xml> --ribbon-json <Ribbon.json>`——写安装设置 → 启动弹窗守卫 `scripts/fwguard.py` → 以 `/cmd INSTALL` 触发官方安装代码 → 部署 Ribbon COM 加载项（复制 DLL + Ribbon.xml/json 到技能目录、写 `HKCU\Software\Microsoft\Office\Access\Addins\MSAccessVCSLib.AddInRibbon` LoadBehavior=3）→ 六项自检（目标 accda、Menu Add-Ins 三键、受信任位置、Ribbon DLL、COM Add-Ins、自包含清单）；
  - **运行位优先**：副本损坏是加载项崩溃的首要根因，**操纵一律优先技能目录运行位 `Microsoft Access Version Control System\`**；安装后跑 `scripts/probe_vcs_api.py` 阶梯探针验证通道可用；
  - **加载项侧操作手册**：安装/卸载 `msaccess-vcs-addin\Wiki\Installation.md`、选项配置 `Options.md`、快速上手 `Quick-Start.md`、MCP 自动化 `MCP-and-Automation.md`（均比本步更详尽，行为疑问先查它们）；
  - **MCP 安装与客户端配置**：
    - 代码位置 `msaccess-vcs-mcp\src\`（可编辑安装进 Python），运行入口 `python -m msaccess_vcs_mcp`；
    - `msaccess-vcs-mcp\.env` 设 `ACCESS_VCS_ADDIN_PATH=<技能目录\Microsoft Access Version Control System\Version Control.accda>`；
    - 客户端配置（`.cursor/mcp.json` 或 `.mcp.json`/`~/.claude.json` 的 `mcpServers` 条目）与端到端验证脚本见 `msaccess-vcs-mcp/README.md` 与第二章；
    - 跑代理生成代码需一次性开启加载项选项 `McpAllowRunVBA`；
  - **调用链与异步语义**：MCP 驱动加载项的底层机制——`Application.Run` → `HandleRibbonCommand` → `clsVersionControl` 的调用链、加载/卸载/配置方式见 `msaccess-vcs-mcp\docs\VBA_INTEGRATION.md`；异步操作的行为语义（`APIAsync` 回调契约、progress/log/error/complete/cancelled 五种消息、取消轮询、同库并发 busy 响应）见 `msaccess-vcs-mcp\docs\VBA_CALLBACK_API.md`——跑大导出/大构建遇到超时或 busy 时按此理解行为并选择等待/取消/重试；
  - **写自动化脚本前抄骨架**：连接与打开库的生命周期管理看 `src\msaccess_vcs_mcp\access_com\connection.py`，`Application.Run` 调加载项/取结果的封装看 `src\msaccess_vcs_mcp\addin_integration.py`，路径校验与只读/写入开关看 `src\msaccess_vcs_mcp\validation.py` 与 `security.py`——真机验证过的写法，比从零写 win32com 稳；
  - **"插件自带 MCP"的机制**：
    - 加载项内置公开 API（`VCS`/`clsVersionControl`，源码见 `msaccess-vcs-addin\Version Control.accda.src\modules\API\clsVersionControl.cls`）与权限门（Options → MCP：`McpAllowImport` 对象导入 / `McpAllowExecuteSQL` 只读 SQL / `McpAllowRunVBA` 任意 VBA，**默认全关**，配置存 `vcs-options.json`）——"被 MCP 操纵的能力"在加载项侧；
    - **MCP 服务器本身**（stdio 协议、工具实现）在 `msaccess-vcs-mcp`（Python：经 COM 连 Access、`Application.Run` 调加载项 API）——加载项=能力提供方（API + 权限门），msaccess-vcs-mcp=协议适配方（MCP 工具 → COM → 加载项 API），两者缺一不可；
    - 权限门启用场景与威胁模型见 `msaccess-vcs-addin\Wiki\MCP-and-Automation.md`。
- **输出物**：VCS 通道可用性报告（安装位=技能目录运行位、加载项版本、MCP 是否注册、探针结果）。

## 阶段 1：需求分析与系统设计

先想清楚做什么，再动数据库。本阶段只产出「设计文档」，不建任何库对象——输入是业务想法，输出是四份文档，它们分别驱动后续的数据建模（表设计说明书）、界面构建（模块清单）与取数设计（业务流程）。

### 1.1 需求与可行性分析

- **任务内容**：明确业务目标、使用者、使用场景、数据规模与部署形态；判断用 Access 是否合适（并发规模、网络形态——C/S、B/S 与文件共享的取舍见 `references/access-architecture-network.md`）。
- **任务要求**：需求的完整性决定后续所有阶段的返工量；数据库设计前的分析工作（先分析什么、问哪些问题、产出什么）见 `references/谈数据库设计前分析工作【Access软件网】.html`。
- **输出物**：需求说明（业务对象、业务流程、角色、约束）。

### 1.2 业务流程梳理

- **任务内容**：把需求落成业务流程——识别业务单据/实体的生命周期（如报销：填报 → 审核 → 付款 → 记账）、每步的参与者与状态流转。
- **任务要求**：画信息流程图（数据从哪来到哪去、谁产生谁消费）——财务系统的信息流程图见 `references/开发财务管理系统.pdf` 第 1 章，盟威教程第 3 章需求设计说明书示范了文字版流程描述；流程图直接决定后续的表清单与查询需求（每一步流转需要什么数据、什么状态字段）。
- **输出物**：业务流程清单 + 信息流程图。

### 1.3 功能与模块设计

- **任务内容**：把业务流程拆成功能模块（如报销系统：基本资料、报销明细、报表、权限），每个模块对应若干窗体/报表。
- **任务要求**：模块划分直接影响表设计与窗体/报表数量——先划模块再设计表；模块与窗体的对应关系见财务系统第 1 章模块设计与盟威教程第 3 章功能设计说明书。**模块边界即高内聚低耦合的第一落点**——每个模块只做一件事、模块间经明确接口协作（窗体 `RecordSource` / 模块公开函数），后续阶段 2 的表、阶段 3 的逻辑层、阶段 4 的界面层都沿此边界展开。
- **输出物**：功能模块清单 + 模块间关系（导航结构草案）。

### 1.4 数据库设计规划（概念 → 逻辑）

- **任务内容**：完成数据库设计规划的前两层——概念建模（实体与关系）与逻辑建模（表设计说明书），为阶段 2 的物理建模提供输入。

数据库设计按「先定数据模型再写代码」的原则分三个层次逐层细化，**本阶段完成前两层，第三层（物理建模）在阶段 2 落实**：

- **概念建模**：识别业务实体与实体间关系（一对一/一对多/多对多），产出 ER 草图——实体即业务对象（客户、订单、订单明细…），关系决定主子表结构；多对多关系需要中间表（阶段 2 落实）。
- **逻辑建模**：把概念模型转成表设计说明书——表清单、每张表的字段/主键/外键/数据类型/必填/默认值、命名规则（表名、字段名、前缀全项目统一）。盟威教程 3.3 节示范了员工代码表、报销明细表的说明书写法。
- **任务要求**：表设计说明书要在建表前完成；命名规则一旦确定全项目统一。
- **输出物**：ER 草图（概念）+ 表设计说明书（逻辑：表清单、字段定义、主键、关系草案）。

## 阶段 2：数据建模与查询设计

本阶段把阶段 1.4 的物理建模落实为真实库对象，并完成取数加工层。输入是表设计说明书与业务流程图，输出是「建好的库 + 查询层 + 第一条版本基线」。先初始化项目并锁定版本基线，再逐层建表、定义关系、设计查询。

### 2.1 项目初始化与版本基线

- **任务内容**：建立项目文件夹与空库，走通第一条纵向切片（一个最小模块的「表 → 窗体 → 输入/输出 → 校验 → 反馈 → 自动化验证」），设计法见 VBA 父技能 `../VBA/SKILL.md`。
- **任务要求**：**库一建好就导出 `.src` 作单一真相源**，后续一切变更走「改 `.src` 文本 → 构建/回写回库 → 漂移核验」闭环。
- **输出物**：项目目录 + 空库 + 第一切片 + `.src` 文本源目录。
- **VCS 融入**：
  - **锁基线**：此时（库刚建好、任何对象都还没改之前）用 `scripts/vcs_drive.py` 的 FullExport 子命令导出全库为 `.src`，随后用 `scripts/vcs_consistency_check.py` 核验漂移为 0——这一步锁定「实时库 ↔ 文本源」基线，输出第一份漂移核验报告，保证后续所有改动都能被跟踪到；
  - **真机演示**：跑 `examples/01_blank_db_to_vcs_loop/run.py`（空库 → 导入模块 → 导出 `.src` → 漂移核验，一条命令跑完全程）；
  - **导出产物格式**——`database.json` 元数据、查询 `.sql` 头注释、模块 `.bas`、BOM/CRLF 编码、文件名净化规则——见 `msaccess-vcs-mcp\docs\EXPORT_FORMATS.md`，与本技能 `references/vcs-text-editing.md` 及对象手册互为补充（文档有出入时以本技能手册为准）；
  - **AI 代理标准工作流**（核心循环图 + 8 种模式：改查询 / 改 VBA / 新建对象 / 批量导出入库 / 团队合并 / 从源码重建 / 重建加载项 / 迭代开发）见 `msaccess-vcs-mcp\docs\AGENT_WORKFLOWS.md`。

### 2.2 建库与表

- **任务内容**：按表设计说明书建库建表（**阶段 1.4 物理建模的落实**：本地表 / 链接表 / 系统表；字段类型、主键、索引、默认值、必填）。
- **任务要求**：
  - **编码铁律**（`.src` UTF-8-BOM、导入 GBK 无 BOM + CRLF、窗体 UTF-16 LE + BOM）见 `references/vba-encoding-rules.md`；
  - **表文本格式**：本地表的 `.xml` 架构、链接表的 `.json` 与 ODBC 连接串一致性规则见 `references/tables-and-relationships.md`；
  - **格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/tbldefs/tOrders.xml`（ID 主键 + 外键索引），照抄结构再改字段名即可；多级（父子）结构的建模参考 `examples/Access BOM Management System/README.md`（BOM 系统完整案例，含 `BOM.accdb`、`treeview.accdb` 可打开对照）；
  - **业务编号**：订单号/流水号字段直接复用 `examples/Access VBA Modules Collection/` 的 `basAutoNumStr`（`AutoNumStr("订单表","订单号",5,"ORD","yyyymmdd")`，用法见其 README）；
  - **建表步骤**：盟威教程第 4 章、财务系统见 `references/开发财务管理系统.pdf` 第 2 章（数据库创建、建表、表间关系）。
- **输出物**：物理表 + 每张表的字段定义核对表。
- **VCS 融入**：每完成一个增量（一张表或一个关系）立即导出核验——`scripts/vcs_drive.py` 的 FullExport 子命令后 `scripts/vcs_consistency_check.py`，确保文本源与实时库始终对齐；改表结构=改文本再回写，而不是在 Access 界面里改完就忘。

### 2.3 表间关系

- **任务内容**：建立表间关系（主外键、级联规则、引用完整性），区分父子（主子）关系与一般关联。
- **任务要求**：
  - **时序**：关系定义要在查询与窗体之前完成，否则查询 JOIN 与主子窗体无从设计；
  - **删除规则**：关系的 `.json` 编辑与删除须连同 `vcs-index.json` 条目一起删（见 `references/vcs-text-editing.md`）；
  - **格式**：关系的 JSON 结构、`Attributes` 引用完整性值表（0=强制/2=松散/16777216=左外/33554432=右外）、复合键写法与建表外键命名规范见 `references/tables-and-relationships.md`；
  - **格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/relations/tCustomerstOrders.json`（`Attributes=0` 强制完整性）与其配套表定义，照抄结构再改字段名即可。
- **输出物**：关系图 + 每对关系的外键字段清单。
- **VCS 融入**：同 2.2——每完成一个关系增量立即 FullExport + 漂移核验，改关系定义=改 `relations/*.json` 文本再回写（删除关系须连同 `vcs-index.json` 条目一起删）。

### 2.4 查询设计与 SQL 编写

SQL 在本技能中出现在三个位置：**查询对象（`.sql` 文本）**、**窗体/报表的 `RecordSource`**、**VBA 中的 DAO/ADO SQL**。本步骤管查询对象，后两处分别在阶段 4/5（RecordSource）与阶段 3（DAO/ADO）落地。

- **任务内容**：按业务取数需求设计查询。
  - **查询类型**：选择查询（筛选/排序/计算/参数）、操作查询（更新/删除/追加）、联合查询、传递查询、子查询；
  - **取数示范**：财务系统的查询设计示范（联合/传递/子查询/日记账查询）见 `references/开发财务管理系统.pdf` 第 3 章；盟威教程第 5 章讲为何建查询与创建方法。
- **任务要求**：
  - **查询必须可更新**（避免 #3027 不可更新错误）——用 `scripts/enumerate_nonupdatable_queries.py` 静态扫描（`--db` 指定目标库，缺省扫 Edonsoft 后端 `Edonsoft Development Framework_x64/Data.accdb`）；
  - **文本格式细则见 `references/queries.md`**——`.bas` 元数据与 `.sql` 恒成对、单一 `Begin Joins` 块规则（多 JOIN 必须写进一个块，否则导入致命错）、`SaveQuerySQL=true` 时 `.bas` 禁止含 `dbMemo "SQL"`、直通查询用 `dbMemo "Connect"`；
  - **格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/queries/qOrders.sql`（格式化 JOIN + ORDER BY）与其 `.bas` 配对，新查询照此模板写；
  - **取数形态参考**：复杂 `RecordSource` 的通用查询接入见 Edonsoft `USysFrmCommonQuery`/`CommonQuery_*`（`Edonsoft Development Framework_x64/edonsoft-docs/03-query.md`）；财务系统的查询设计示范（联合/传递/子查询/日记账查询）见 `references/开发财务管理系统.pdf` 第 3 章；盟威教程第 5 章讲为何建查询与创建方法；
  - **查询源文件与内部存储**
  - **SQL 语法权威参考**：Access SQL 保留字/语句/谓词大全（48 篇）见 `..\VBA-Docs\access\Concepts\Structured-Query-Language\`（如 `create-and-delete-tables-and-indexes-using-access-sql`、`define-relationships-between-tables-using-access-sql`、`all-distinct-distinctrow-top-predicates-microsoft-access-sql`）；
  - **条件表达式参考**：WHERE 条件/域聚合函数/日期文本条件（19 篇）见 `..\VBA-Docs\access\Concepts\Criteria-Expressions\`（含变量/控件参与 SQL 与日期条件的官方写法）；
  - **查询执行 API 单页**：打开/执行查询用 `..\VBA-Docs\api\Access.DoCmd.OpenQuery.md` 与 `Access.DoCmd.RunSQL.md`；VBA 内取数记录集用 `Application.CurrentDb`（`..\VBA-Docs\api\Access.Application.CurrentDb.md`）配合 DAO（对象模型见 `references/vba-database-programming.md`）；：`.sql` 文本形态与手写/校验规则见 `msaccess-vcs-addin\Wiki\Query-Source-Files.md`；MSysQueries 字段、Design View vs SQL View、`LoadFromText`/`SaveAsText` 不对称性与解析器不变量见 `msaccess-vcs-addin\docs\access-query-storage.md`（理解 `.bas`/`.sql` 恒成对与回写形态变化的权威参考）；导出文件元数据头（`-- Query:`/`-- Type:`/`-- Exported:`）与文件名净化见 `msaccess-vcs-mcp\docs\EXPORT_FORMATS.md`。
- **输出物**：查询清单（名称、SQL、用途、可更新性结论）。
- **VCS 融入**：改查询=编辑 `.sql` 后回写，构建后核验漂移 0；用 `scripts/enumerate_nonupdatable_queries.py` 的扫描结果作为本环节的质量门。

### 2.5 性能设计（索引、查询与窗体加载）

- **任务内容**：在数据层落实性能设计——索引策略、查询写法约束、窗体/子窗体加载约束，为阶段 7 的性能断言打底。
- **任务要求**：
  - **索引**：主键与高频外键/筛选/排序字段建索引，避免对函数/表达式列建索引（会失效）；
  - **查询**：能下推到查询的取数不要用 VBA 循环；避免 `WHERE` 内对字段做函数包裹（`DateDiff`/`Format` 使索引失效）；避免 `SELECT *`；慢查询先跑 `scripts/enumerate_nonupdatable_queries.py` 定位；
  - **窗体加载**：列表窗体用连续窗体减少逐行重算；主子窗体嵌套层级与记录源数量是加载慢主因（子窗体越少越好）；大表先用筛选查询缩小记录集再绑窗体；启动窗体只做必要初始化；
  - **优化顺序**「表/索引/查询 → 窗体/子窗体 → VBA/循环」；数据量超单文件共享承受力时，前后端分离/后端升迁是终局手段（见 8.1）。
- **输出物**：索引清单 + 慢查询/大窗体清单（含优化结论）。
- **VCS 融入**：索引/查询改动走 2.1 闭环——改 `tbldefs`/`queries` 文本 → 回写 → 漂移核验。

## 阶段 3：通用代码库（VBA 逻辑层）

VBA 在本技能中的编写落点有三处：**标准模块**（本阶段 3.1，可复用的业务函数）、**类模块**（本阶段 3.2，有状态的对象封装，即面向对象建模在 Access 的落地）、**窗体/报表代码后置**（阶段 4/5，事件驱动代码，随窗体报表一起走 `.cls`）。宏（本阶段 3.3）承载启动与自动化动作。输入是功能模块清单，输出是可供窗体/报表/查询调用的逻辑层。

### 3.1 标准模块

- **任务内容**：把可复用的业务逻辑沉淀为标准模块（自动编号、字段校验、ADO 执行、导出等），供窗体/报表/查询调用。
- **任务要求**：
  - **优先复用而非新写**：`examples/Access VBA Modules Collection/` 提供：
    - 自动编号（`basAutoNumStr`）、字段校验（`ClsFieldValidator` + `M_ValidationUI`）、ADO 操作（`ADOExecute.cls`）
    - Excel/PPT/HTML 导出（`modExportToExcel`/`modExportToPPT`/`modHTMLExport`/`modPasteDataToExcel`/`basExportChart`）、VBE 工具（`modVBETools`，批量注入标准错误处理）
    - 每个模块的用途、接口与用法见其 README 与 `wiki/`；财务系统第 5 章示范了模块与通用/专用函数的分层；需要新写时用 `templates/standard-module-stub.bas` 起步；
  - **生产级模块分层范本**：`msaccess-vcs-addin\Version Control.accda.src\modules\` 按 API/Core/Infrastructure/Install/Integration/Interfaces/Lib/Tests/Utility 十目录组织——真实加载项工程的模块划分与错误处理方式，写自己的通用模块库前可对照；
  - **公开 API 面范本**：Edonsoft 把 60 个公开函数按 7 类组织（编号/SQL/窗体记录/杂项/IO/UI 交互/安全，全部 `Application.Run` 可调用，见 `Edonsoft Development Framework_x64/edonsoft-docs/11-api.md` 与 `api/` 目录）；代表性落点：`AutoNumStr`（流水号）、`ReadRecord`/`WriteRecord`/`DeleteRecord`（记录读写）、`JsonConverter`（JSON）、操作日志 `WriteLog`/`AuditData`（审计字段，见 `13-architecture-source.md` 第 8 节第 9/18 行）；
  - **框架可见源码模块**
  - **记录集参考**：DAO 记录集操作官方做法（增删改查、遍历、一对多关系）见 `..\VBA-Docs\access\Concepts\Data-Access-Objects\`（27 篇）与 `ActiveX-Data-Objects\`（ADO）；
  - **对象模型 API 单页**：Access 应用级对象（`Application.CurrentDb`/`CurrentProject`/`CurrentUser`、`DoCmd.RunSQL`/`OpenForm`、`Module.AddFromString`）按名查 `..\VBA-Docs\api\Access.*.md`（总索引 `api\Access.md`、`api\overview\Access.md`）；
  - **VBA 语言总索引**：关键字/函数/语句/对象/常量见 `..\VBA-Docs\Language\Reference\`（`keywords-visual-basic-for-applications`/`functions-visual-basic-for-applications`/`statements`/`objects-visual-basic-for-applications`/`constants-visual-basic-for-applications`）；（可直接读的 VBA 实现，清单见 0.2 与 `framework-overview.md` 第 4 节）：
    - `basCoreReference`（启动引用修复）、`basZip`（压缩）、`basButton`/`basFormFontTools`（UI 工具）
    - `clsLog`（操作日志类）、`JsonConverter`（JSON）、`Module_DatePicker`/`Module_YearMonthPicker`（日期控件，与 `examples/Access DatePicker` 同源）
    - 其中 `Module_DatePicker`/`Module_YearMonthPicker` 与 `basCoreReference` 可直接移植进业务库。
- **输出物**：可调用的标准模块（函数/过程 + 用途注释）。
- **VCS 融入**：**导入走自动化，不手工粘贴**——用 `scripts/rebuild_module_from_src.py` 原子重建（文件转 GBK 无 BOM + CRLF → 按 `Attribute VB_Name` 原子替换 → `Import` → 读回逐行比对），或经 VCS MCP 的 `vcs_import_object` 回写；改码在 `.src` 文本上进行再构建，全程零手工粘贴。

### 3.2 类模块（面向对象建模）

- **任务内容**：封装有状态的对象（如字段校验器、搜索组合框、事件处理器），用属性/方法/事件表达业务对象的行为——实体逻辑不散落在窗体事件里，而是收进类。**类模块是面向对象建模与高内聚低耦合的核心手段**：状态与行为封装在一个类内，外部只经属性/方法访问，不直接触碰内部实现。
- **任务要求**：
  - **类模块坑与格式**：`type` 退化是常见坑（导入后变成空白/错误类型）——编码与导入规则见 `references/vba-encoding-rules.md`；**VBA 文本格式细则见 `references/vba.md`**——`.cls` 不得含 `VERSION 1.0 CLASS` 头与 `Attribute VB_Name`（VCS 导入自行处理，含了必报错）、64 位 API 声明用 `#If VBA7 Then` 条件编译、超链接列写入规则（`#`-分隔三元组、地址必须绝对、文件名禁含 `#`）；模板见 `templates/class-module-stub.cls`；
  - **格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/modules/modNavigation.bas`（标准模块：`Attribute VB_Name` + 共享导航函数，`forms/*.cls` 为代码后置格式样板），写模块/类模块前先照此核对；
  - **现成类模块示例**：`examples/Access VBA Modules Collection/` 的 `ClsFieldValidator.cls`（字段校验器，属性/方法/事件齐全）与配套 `M_ValidationUI.bas`（校验结果联动 UI），可直接导入改写；
  - **生产级类模块范本**：`msaccess-vcs-addin\Version Control.accda.src\modules\` 下的 `clsVersionControl`（公开 API 实现）、`clsQueryComposer`（查询解析/合成）、`clsSourceParser`/`clsConditionalFormat`/`modLoadSaveText`（导出导入管线）——写有状态类或解析器类前对照其接口设计与错误处理。
- **输出物**：可实例化的类模块 + 使用示例。
  - **类模块与错误处理官方参考**：类模块编程（`program-with-class-modules`、自定义属性与方法）、运行时错误处理（`elements-of-run-time-error-handling`/`error-trapping`）、`DoCmd` 对象与宏动作对照见 `..\VBA-Docs\access\Concepts\Error-Codes\`；VBA 事件模型与语言概念见 `..\VBA-Docs\Language\Concepts\` 与 `..\VBA-Docs\Language\Reference\events-visual-basic-for-applications.md`。
  - **类模块工程组织**（属性/方法/事件 + 接口设计 + 错误处理）对照 `msaccess-vcs-addin\Version Control.accda.src\modules\` 的 `clsVersionControl`/`clsQueryComposer`；代码模块动态操作（`Module.AddFromString`/`AddFromFile`）按名查 `..\VBA-Docs\api\Access.Module.*.md`；

- **VCS 融入**：同 3.1——`.cls` 的导入/重建走 `scripts/rebuild_module_from_src.py` 原子重建（VCS 导入自行处理 `VERSION 1.0 CLASS` 与 `Attribute VB_Name`，文本中不得含），改码在 `.src` 文本上进行再构建，全程零手工粘贴。

### 3.3 宏

- **任务内容**：用宏承载启动与自动化动作（AutoExec 在打开库时自动运行）。
- **任务要求**：宏在 `.src` 中是 `.bas` 文本（Access 宏动作，非 VBA）；AutoExec 模板见 `templates/autoexec-macro.macro`（无 BOM）；宏/共享图像/导入导出规格的对象格式见 `references/other-objects.md`（宏的 `Version=196611` 与 `Begin...End` 动作块、共享图像 `.json`+图像文件配对、imexspec 列映射结构）；用代码运行宏按名查 `..\VBA-Docs\api\Access.DoCmd.RunMacro.md`。
- **输出物**：AutoExec 等宏 + 触发说明。
- **VCS 融入**：模块/类模块/宏的任何变更都走「改 `.src` 文本 → `scripts/rebuild_module_from_src.py` 或 VCS 构建回写 → 核验」；改完用 `Application.Run` 或脚本调用验证函数可用。

## 阶段 4：窗体与界面（交互层）

把阶段 3 的逻辑层接到用户面前。输入是功能模块清单与查询层，输出是完整的导航界面。按「窗体 → 控件 → 菜单导航」三层推进：先建窗体承载数据与操作，再在窗体上落地控件，最后把窗体挂进导航。

### 4.1 窗体

- **任务内容**：按模块设计构建用户界面——列表窗体、编辑窗体、主子表窗体、主控窗体；窗体上事件驱动的 VBA 代码（按钮点击、值变更、打开关闭）写在各窗体的代码后置（`.cls`）。窗体坚持**薄窗体厚逻辑**——事件里只做接线与交互，业务逻辑调阶段 3 的模块/类，不把 SQL 与业务规则写进事件过程（交互/逻辑/取数各归其层的关注点分离）。
- **任务要求**：
  - **窗体文本格式细则见 `references/forms.md`**——8 条关键规则（TabIndex 连续、VBA 事件必须在 `.bas` 接线否则静默不执行、禁改二进制块、`\"` 转义、宽度重算等）、窗体骨架结构（`Version=20` → 单一容器 → 默认样式 → 节）、组合框三坑（`RowSourceTypeInt` 配对、`BoundColumn` 0-based、`AllowValueListEdits`）、子窗体宽度公式；
  - **格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/forms/fCustomerList.bas`+`.cls`（列表窗骨架与事件接线）、`fOrderDetail.bas`（含 `fOrderItemsSub` 主子窗体嵌套），写窗体前先照此核对；
  - **拆分文件规则**：`.bas`/`.cls` 拆分文件形态（`SplitLayoutFromVBA` 为真时的布局/代码分离）见 `msaccess-vcs-addin\Wiki\Split-Files.md`；
  - **窗体完整示范**：财务系统第 4 章（主窗体、日记账输入、分类账管理、显示窗体、试算、报表窗体、主控窗体、欢迎窗体）与盟威教程第 6 章（自动创建窗体、数据模块创建向导、配置报销明细/员工编码窗体）；
  - **优先用内建生成器**：Edonsoft 代码生成器（`USysFrmCodeGeneratorPro`，见 `Edonsoft Development Framework_x64/edonsoft-docs/02-codegen.md`）按表结构生成主窗体/列表/编辑/主子表窗体与菜单项。
- **输出物**：窗体集合（名称、用途、RecordSource、按钮行为清单）。
  - **窗体事件与记录操作官方示范**：窗体级事件/过滤/验证/多实例（`apply-a-filter-when-opening-a-form-or-report`、`perform-simple-data-validation-checks-when-editing-a-record-in-a-form`、`create-multiple-instances-of-a-form`、`use-user-input-to-build-filter-criteria`）见 `..\VBA-Docs\access\Concepts\Forms\` 与 `Forms-Design\`；窗体对象模型单页按名查 `..\VBA-Docs\api\Access.Form.*.md`（252 页，如 `Access.Form.Requery.md`/`Access.Form.Dirty.md`）与 `..\VBA-Docs\api\Access.DoCmd.OpenForm.md`（打开/关闭/条件传参）；

- **VCS 融入**：窗体改动走 2.1 统一闭环（改 `.src` 文本 → 构建/回写回库 → 漂移核验），本步差异动作——布局/属性/事件代码改动前先 FullExport 拿到最新 `.bas`/`.cls` 文本，在文本上改再回写；删除窗体时必须同时删配对文件（`.bas`/`.cls`/`.json`）与 `vcs-index.json` 条目（规则见 `references/vcs-text-editing.md`）。

### 4.2 控件

- **任务内容**：在窗体上落地具体控件。
- **任务要求**：
  - **日期选择**：用 `examples/Access DatePicker/README.md`（`Module_DatePicker.bas` / `Module_YearMonthPicker.bas`，含 `DatePicker.accdb` 演示库与 `使用说明.md`）；
  - **输入校验与数据存取**：`examples/Access VBA Modules Collection/` 的 `ClsFieldValidator.cls`（字段规则校验）+ `M_ValidationUI.bas`（校验结果高亮联动）、`ADOExecute.cls`（供控件/窗体做数据存取）；
  - **Tab 控件只能 COM 构建**（细节见 `references/vba-com-automation.md` 铁律 5）；
  - **搜索组合框与单据选择**：Edonsoft 公开 API——`CreateSearchCombo`（实现类 `clsSearchCombo`，见 `Edonsoft Development Framework_x64/edonsoft-docs/04-search-combo.md`）与单据选择窗体 `USysFrmBillSelect`/`BillSelect_ShowForSubform`（`Edonsoft Development Framework_x64/edonsoft-docs/10-bill-select.md`）；
  - **进度条**：参考 Edonsoft `USysFrmProcessBar`/`USysFrmProgressFlat/Dual/Detail`（基础/详细/双层/扁平，百分比+剩余时间+速度，见 `13-architecture-source.md` 第 8 节第 16 行）；
  - **图片与条件格式**：图片增删/替换/OLE 转换见 `references/images.md`；条件格式（解码→删除二进制块→VBA `FormatConditions` 重建）见 `references/conditional-formatting.md`，`ConditionalFormat`/`ConditionalFormat14` 二进制格式的权威逆向规范（含未验证遗留块）见 `msaccess-vcs-addin\docs\access-conditional-format.md`，解码器实现见 `Version Control.accda.src\modules\Core\clsConditionalFormat.cls`；
  - **界面资源**：Edonsoft `assets/`（css/js/images/Theme，内嵌网页与仪表盘的前端资源）与盟威 `Images/`（登录界面与主题图片：login_classic/login_standard/head.jpg/progress.jpg）；Access 内嵌网页/仪表盘用的前端库（Bootstrap / jQuery / ECharts / Font Awesome，清单见 `assets/manifest.json`、说明见 `assets/ASSETS.md`）在 `assets/` 下。
- **输出物**：窗体上的控件 + 事件代码。
  - **控件官方示范**：组合框/列表框/选项卡/子窗体官方做法（`addallto-a-combo-box-or-list-box`、`synchronize-two-combo-boxes-on-a-form`、`refer-to-tab-control-objects-in-vba`、`hide-a-subform-if-the-main-form-contains-no-records`、`call-procedures-in-a-subform-or-subreport`）见 `..\VBA-Docs\access\Concepts\Controls\`（7 篇）。
  - **控件对象模型单页**：`Control` 及各控件类型的属性/方法/事件按名查 `..\VBA-Docs\api\Access.*.md`（每类 100–180 页：`Access.ComboBox.*` 178 页、`Access.TextBox.*` 163 页、`Access.ListBox.*` 145 页、`Access.CommandButton.*` 143 页等）；


### 4.3 菜单、导航与主界面

- **任务内容**：把业务窗体挂进导航（主界面菜单 / 启动菜单）。
- **任务要求**：
  - **菜单形态**：盟威教程 6.7–6.8 节示范了查阅列表与启动菜单；Edonsoft 的动态菜单由数据配置生成（`USysFrmMenuSettings` 配置界面、`FArguments` 动作参数格式，见 `Edonsoft Development Framework_x64/edonsoft-docs/13-architecture-source.md` 第 5 节）；
  - **待办列表（可直接借鉴）**：`USysFrmToDoList` + `AddReminder`/`GetUnreadCount`/`MarkReminderRead`/`DeleteReminder`——数据源=业务表按状态/负责人筛选 + 未读徽标 + `FArguments` 业务跳转（见 `Edonsoft Development Framework_x64/edonsoft-docs/08-todo.md`）；
  - **仪表盘（可直接借鉴）**：`USysFrmDashboard` + `Dashboard_AddKPIFromSQL`/`Dashboard_AddChartFromSQL`/`Dashboard_Render`——WebBrowser 渲染 KPI 卡片与 ECharts 图表、12 栅格（见 `Edonsoft Development Framework_x64/edonsoft-docs/09-dashboard.md`）；
  - 待办 = 业务表按状态/负责人筛选 + 主界面链接；仪表盘 = 汇总查询/图表 + 主界面子窗体——主界面与菜单无需从零做。
- **输出物**：可导航的菜单结构 + 主界面。
- **VCS 融入**：窗体层差异动作见 4.1；本步差异动作——菜单/导航数据（`FArguments` 配置、导航窗体、启动菜单）的改动同样走 2.1 统一闭环，在文本源上改再构建回库。

## 阶段 5：报表设计（输出层）

把业务数据变成可打印的输出。输入是查询层与窗体结构，输出是报表集。按「静态 → 动态 → 切换打印」推进：先做固定格式报表，再做按条件取数的动态报表，最后组织成可切换面板接入主界面。

### 5.1 静态报表

- **任务内容**：制作固定格式的输出报表（标题、字段、格式）。
- **任务要求**：静态报表的制作步骤见盟威教程 7.1–7.2 节；财务系统各报表（分类账、日记账、利润表、资产负债表、财务指标）见 `references/开发财务管理系统.pdf` 4.4–4.12 节；**需要输出到 Excel/PPT/HTML 时直接用导出积木**——`examples/Access VBA Modules Collection/` 的 `modExportToExcel`（导出报表到 Excel）、`basExportChart`（图表导出）、`modExportToPPT`（PPT 汇报输出）、`modHTMLExport`（HTML 输出）、`modPasteDataToExcel`（数据粘贴并格式化到 Excel 模板）。
- **输出物**：静态报表集（名称、数据源、用途）。
  - **报表与打印官方参考**：报表概念（分组/汇总/条件）见 `..\VBA-Docs\access\Concepts\Reports\`，打印与分页（页面设置、打印前/后事件）见 `..\VBA-Docs\access\Concepts\Printing\`（5 篇）。
  - **报表对象模型单页**：`Report` 及报表节/分组对象属性方法按名查 `..\VBA-Docs\api\Access.Report.*.md`（178 页）；打印入口用 `..\VBA-Docs\api\Access.DoCmd.OpenReport.md`；

- **VCS 融入**：同 5.3（报表无特有差异动作）。

### 5.2 动态报表

- **任务内容**：按参数/筛选条件动态取数的报表。
- **任务要求**：动态报表的制作见盟威教程 7.3 节；报表的排序分组、汇总与格式化规则见 `references/reports.md`（报表节关键字、`BreakLevel` 排序分组与多级分组、节宽 vs 页宽约束含纸张宽度表、Label 属性限制、常见结构错误清单）；**格式活样例**：`examples/CustomerOrders/CustomerOrders.accdb.src/reports/rOrderReport.bas`（带 `OrderID` 过滤条件的分组报表），新报表照此模板改数据源与分组即可。
- **输出物**：动态报表集（名称、筛选参数、用途）。
- **VCS 融入**：同 5.3（报表无特有差异动作）。

### 5.3 报表切换与打印

- **任务内容**：把多个报表组织成可切换的面板，接入主界面。
- **任务要求**：动态报表切换面板见盟威教程 7.4 节；报表在窗体/主控窗体中的引用方式见财务系统 4.13 节。
- **输出物**：报表集 + 切换面板 + 打印入口。
- **VCS 融入**：报表改动走 2.1 统一闭环（改 `.src` → 回写 → 核验），无报表特有差异动作。

## 阶段 6：AI 能力接入

在系统上叠加 AI 能力，分两个方向：业务系统内嵌聊天问答（面向终端用户），以及用 AI 编码代理在文本源上改库（面向开发本身，即本技能的工作方式）。

### 6.1 内接大模型（业务系统内聊天问答）

- **任务内容**：让业务系统内嵌 AI 对话（多模型、连续对话、库表分析、SQL 建议）。
- **任务要求**：用 `examples/AccessAI/README.md`（升级版「Access LLM Toolkit」：`AI.accdb` 完整实现 + `JsonConverter.bas` + `Module_Markdown.bas` 等源码，支持流式输出、对话历史持久化、Access SQL 助手、TXT/CSV/Word/Excel/PDF 文档问答、DPAPI 加密 Key）——把 AI 窗体与调用代码导入业务库即可；Edonsoft 框架自带 AccessAI（`basAI`、`FS_AIWeb`、`FS_ChatHistory`，见 `Edonsoft Development Framework_x64/edonsoft-docs/13-architecture-source.md` 第 9 节）。
- **输出物**：可对话的 AI 窗体 + 调用入口。
- **VCS 融入**：导入 AI 窗体/模块后立即 FullExport + 漂移核验，把 AI 能力纳入文本源管理（同 2.1 闭环）。

### 6.2 AI 编码代理与文本化改库

本技能自身的工作方式就是 AI 编码代理在文本源上改库：**导出 `.src` → 在文本源上改 → 构建回库**。AI 能创建/编辑的对象与人工一致（表/关系/查询/窗体含主子窗体/报表含分组汇总/模块与事件代码/库属性与启动配置），改库的唯一入口是文本源，绝不让 AI 直接操作 Access 界面点按。

- **任务内容**：用 AI 编码代理（本技能自身）在文本源上改库。
- **任务要求**：
  - **从 0 到 1 生成新库**：用 `examples/CustomerOrders/prompt.txt` 作 prompt 模板——保留其句式结构（启动屏 → 列表/编辑窗 → 主子窗体 → 预览报表 → 汇总行），替换业务对象、字段与报表要求即可；生成的库以 `examples/CustomerOrders/CustomerOrders.accdb.src/` 为格式基准（该样例 3 表 4 查询 6 窗体 1 报表，一个 prompt 生成，含 `CustomerOrders.accdb` 可运行成品）。
  - **改现有库**：先读 `.src`（编码 UTF-8-BOM + CRLF，`references/vba-encoding-rules.md`），改文本的硬规则（对象配对、编码换行、安全删除）见 `references/vcs-text-editing.md` 与其中第六节的映射表；**改哪种对象先读对应手册、再到示例核对格式**——样例是 9 份对象手册的活样例，两条对照链（手册↔示例↔实时库）保证改动一次成型。
  - **门禁（硬性）**：AI 修改后的文本源**未通过门禁不得进入构建**——① 漂移核验 0 项（`scripts/vcs_consistency_check.py`，基线定义见 2.1、构建成功判别见 7.4）；② 编译门禁 `CompileVBA` 通过（`scripts/vcs_drive.py CompileVBA --db <库>`，手段见 7.2 手段 D）；③ 改动对象逐项 `ExportObject` 复核产物落盘。三条全绿才允许构建回库；否则回到文本源修正后重跑门禁。
  - **安全**：AI 编码代理应在隔离环境运行，不应拥有宿主机的不受限访问权限。
- **输出物**：AI 修改后的 `.src` 文本源（可审计的 diff）+ 门禁通过记录。
- **VCS 融入**：
  - **每轮改动走 2.1 统一闭环**（改 `.src` → 构建回库），本步差异动作是**运行断言**——`Application.Run` 断言改动生效，对齐「改了什么 → 构建是否成功 → 运行是否生效」三段；这个循环的最小演示见 `examples/02_edit_src_reimport/run.py`（改 `.src` 一处 → 原子重导入 → 断言生效），跑通它即掌握了 AI 改库的基本单元；
  - **AI 代理八种操作模式**（修改查询含编译失败停止约定、增改 VBA 保留 `Attribute VB_Name`/`Option Explicit`、新建对象、批量导出入库、团队 pull 合并、从源码重建发布库、重建加载项、迭代开发）逐条步骤与判定见 `msaccess-vcs-mcp\docs\AGENT_WORKFLOWS.md`（把「何时导出、何时合并、何时编译、何时测试」写成可照做的清单）；
  - **实现 `vcs_*` 工具等价能力**：直接读 `msaccess-vcs-mcp\src\msaccess_vcs_mcp\tools.py` 里对应工具的实现——导出/导入/重建/编译/执行 SQL/调用 VBA 的参数解析、返回结构（`success`/`log_path`/`error`）、失败分支处理都有真机验证过的写法可抄；
  - **保存即导出（可选增强）**：让开发期"保存即自动导出 `.src`"，参考 `msaccess-vcs-addin\Wiki\Export-on-Save-Hook.md`（`modExportOnSaveHook` 的 `OnSave` 钩子机制，源码见 `msaccess-vcs-addin\Hook\`）。

## 阶段 7：自动化测试与验证（质量门）

让系统行为可被自动验证，并给「改文本源 → 构建」设置硬性质量门。输入是已构建的库，输出是验证脚本与断言报告。按「可测性设计 → 自动化断言 → 构建成功判别」推进。

### 7.1 可测性设计

- **任务内容**：把可测逻辑写成不依赖界面的纯函数（TDD 分层与 COM 自动化 SOP 见 `references/vba-com-automation.md`），界面行为留待 COM 脚本驱动断言。
- **任务要求**：验证全用脚本零手动——定位工程用 `scripts/fwguard.py`（按 `FileName` 遍历 `VBProjects`）、模板验证 `scripts/probe_templates.py`、开工探测 `scripts/diagnose_vba_project.py` 与 `scripts/probe_vba_project.py`。
- **输出物**：可测函数清单 + 测试脚本骨架。

### 7.2 自动化测试手段（弹窗/报错/调试/VBE，全代码化）

- **任务内容**：按被测行为选择对应手段，全程零手动。完整方法见 `references/vba-com-automation.md`（本步即手段 A–F 的展开单点）；调试日志模板 `templates/modDebugLog.bas`。
- **测试资源（MCP 与加载项侧）**：
  - MCP 工具自身三层测试（venv 单元测试、集成测试、端到端六步：导出→改→合并→验证→重建→比对）与常见失败场景（加载项未装、权限拒绝、无效路径、fast save）见 `msaccess-vcs-mcp\docs\TESTING.md`——其中「默认安装位 `%AppData%\MSAccessVCS\`」与本技能 0.3 自包含原则不同，以本技能为准；
  - 技能自建的 MCP 回归用 `tests\vcs_mcp_verify.py`；加载项侧测试组织（`vcs_run_tests` 与测试类）见 `msaccess-vcs-addin\Wiki\Testing.md` 与 `Regression-Testing.md`。
  - **VBA 错误消息与 Access 错误码权威清单**：VBA 运行时错误号与消息见 `..\VBA-Docs\Language\Reference\error-messages.md`；Access 数据访问错误（含中文乱码导致 SQL 3075 一类）与错误处理要素见 `..\VBA-Docs\access\Concepts\Error-Codes\`（9 篇，`elements-of-run-time-error-handling`/`error-trapping`）。
  - **错误码对象模型入口**：`Access.Application.AccessError`（`..\VBA-Docs\api\Access.Application.AccessError.md`）按错误号取 Access 错误消息文本；
- **手段 F——借鉴 MCP 源码的成熟做法**：同一类自动化问题，先看 `msaccess-vcs-mcp\src\` 是怎么解决的再自己写——并发串行化（同库多调用排队、busy 判定）看 `access_gate.py` 与 `operation_manager.py`；COM 连接意外断开的分类与恢复（正常关闭 vs 崩溃）看 `com_recovery.py`；隔离 VBA 代码执行并带超时（防止死循环挂死整条流水线）看 `vba_worker_manager.py`；只读审查/写入开关看 `security.py` 与 `validation.py`。
- **手段 A——消息框捕获与自动点击**：`scripts/fwguard.py` 后台轮询桌面 `#32770` 对话框，`GetWindowText` 读标题+正文 → 按文本首字符策略自动点按钮（含「确/结/束/是/好/继/续/试/助/O/o/Y/y」点确定/是/重试；含「另存/保存/覆盖/替换/删除/确认修改」等危险字样一律点取消/否）→ `fwguard.dialogs_seen()` 回读全部捕获文本供断言。NetUI 对话框（`NUIDialog`/`NetUIHWND`，如 Build Name Conflict）按钮非标准 Win32 子窗口，改用键盘 `VK_RETURN` 触发默认按钮。
- **手段 B——代码报错捕获**：Python 侧 `silent_execute`（`try/except` 捕获 `com_error` 的 error_number/message）；VBA 未处理错误弹的「Microsoft Visual Basic」错误框文本（含错误号与描述）由 fwguard 捕获；需定位出错行时用探针包装（`On Error` 分支调 `LogErr` 记录 `Err.Number/Description/Source/Erl` 到日志文件）。
- **手段 C——代码调试信息捕获**：VBA 侧用 `templates/modDebugLog.bas` 的 `LogPrint`/`LogTimer`/`Watch` 输出到 `.\vba_debug.log`（等价 `Debug.Print` 的文件版，`#Const DEBUG_MODE` 控制生产开关；自动化场景禁止 `MsgBox`，`Debug.Assert` 仅限人工）→ Python 侧回读日志文件断言。
- **手段 D——操纵 VBE**：定位工程用 `fwguard.get_proj(app, path)` 按 `FileName` 遍历（**禁 `VBProjects(1)` 下标**——accda 打开后 VBA 工程延迟加载会越界）；读源码用 `VBComponent.CodeModule` 的 `CountOfLines`/`Lines`；改代码用 `InsertLines`/`ReplaceLine`/`DeleteLines`；改后必须编译门禁（`CompileVBA`）+ 导出复核。
- **手段 E——修改 VBA 代码**：小改走 VBE CodeModule 原位替换；整模块重建走 `scripts/rebuild_module_from_src.py`（文件转 GBK 无 BOM + CRLF 原子重建，符合编码铁律）；改后重编译 + 漂移核验 0 项。
- **任务要求**：测试用临时副本，不拿正式数据；最小用例集含正常/空值/重复/边界；被测试代码路径不出现 `MsgBox`（用日志替代）；禁止 SendKeys/坐标点击依赖时序的手段。
- **输出物**：弹窗捕获日志、错误/调试日志文件、断言报告。

### 7.3 自动化断言

- **任务内容**：用 COM 脚本驱动界面/逻辑断言（打开窗体 → 操作 → 断言结果），技能自身完整性冒烟见 `tests/test_skill_integrity.py`（如何运行、检查什么见 `tests/README.md`）。
- **任务要求**：硬规则全集在 `references/iron-laws.md`；COM 对象销毁按「进程快照区分本次新进程 → finally 逆序释放 → 只杀本次新进程」规范（借鉴 `references/vba-com-automation.md` 与 VBA 技能 com-testing 方法论），禁止无差别 `taskkill` 误杀用户 Office——MCP 的 `access_com\instance_registry.py` 正是"登记本服务创建的 Access 进程、只清自己拉的"这一模式的工程实现，写清理逻辑前先看它。
- **输出物**：验证脚本 + 断言结果报告。

### 7.4 构建成功判别与回归收口

- **任务内容**：
  - **构建成功硬性判别（同时满足）**：Build 日志无错、备份库已生成、库大小明显不同、`scripts/vcs_consistency_check.py` 漂移核验 0 项（判别细节见 `references/vcs-three-channels.md` 的「怎么判别构建成功」一节；增量合并构建 MergeBuild 的前置与语义见 `msaccess-vcs-addin\Wiki\Merge-Build.md`）；
  - **改了加载项自身时的回归**：走 `vcs_rebuild_addin` 重建 + `vcs_run_tests` 跑加载项自带测试套件（`--filter clsTestInstall` 等，headless 输出、进度/完成判定、超时恢复方法见 `msaccess-vcs-mcp\docs\AGENT_WORKFLOWS.md` 6b/6c 节；测试套件类模块源码见 `msaccess-vcs-addin\Version Control.accda.src\modules\Tests\`，写自己的自动化测试类可对照其组织）；
  - **长操作进度与审计**：借鉴 MCP 成熟实现——`src\msaccess_vcs_mcp\usage_logging.py`（结构化 JSONL 调用日志，`ACCESS_VCS_ENABLE_LOGGING` 开关）与 `callback_server.py`/`operation_manager.py`（异步操作跟踪、超时、取消），写自己的长任务脚本时照此组织。
- **任务要求**：任何一次「改文本源 → 构建」都以这四条件收口；回归发现问题回到对应阶段修正，修正后重跑全部验证。**全技能回归收口**：`python tests/test_skill_integrity.py`（完整性 55 项）→ `python tests/vcs_com_verify.py --db <宿主库>`（COM 层）→ `python tests/vcs_mcp_verify.py --db <宿主库>`（MCP 端到端）三连全绿；重构建走 `python scripts/rebuild_and_install.py`（修复门禁 → 构建 → 代码化安装 → 探针回归，一键）后同样三连验证。
- **输出物**：构建成功判定记录 + 回归结论。

## 阶段 8：部署与运维（交付）

把验证过的系统交付使用。输入是完整库与验证报告，输出是可部署的前端副本、权限配置、启动配置与版本快照。按「数据与程序分离 → 权限 → 发布更新 → 收尾核验」推进。

### 8.1 前后端分离与链接表

- **任务内容**：把后端数据与前端程序分离（后端 `Data.accdb` 放共享位置，前端通过链接表访问），便于多人共享与升级。
- **任务要求**：
  - **链接表的连接与批量重链**：Edonsoft 用 `LinkedTableMgr_*`（`USysFrmLinkedTableManager`，见 `Edonsoft Development Framework_x64/edonsoft-docs/06-linked-manager.md`），多后端连接串统一封装见 `BuildAccessAdoConnectionString`（`05-backend-deploy.md`），升迁 SQL Server/MySQL 见 `USysFrmUpsize`/`Upsize_*`（`05-backend-deploy.md` 与 `07-upsize.md`）；
  - **连接与重链语义**（DSN 一致性、连接串缓存）见 `msaccess-vcs-addin\Wiki\Connections.md`；
  - **多后端配置化切换参考**：Edonsoft `appsettings.json`（`ActiveBackend` + `Backends[]` 数组，每项含 `Type`/`DbType`/`AccessPath`/`Server`/`DbConnect`/`DriverName`，启动时按 Active 解析连接串）；
  - **盟威**：前后端分离由平台自身保证（`Data.accdb` 后台数据库，教程第 4.3 节「链接后台数据库」）。
  - **架构意义**：前端=程序（窗体/报表/代码），后端=数据（表）——前端可随时从 `.src` 重建、后端独立备份，这是分层架构在物理部署层的体现；与 4.2/4.3 界面前端资源合起来构成「程序-数据」「界面-逻辑」两层前后端分离。
- **输出物**：可部署的前端副本 + 链接表配置。

### 8.2 用户权限

- **任务内容**：配置登录账号、角色、菜单/功能权限。
- **任务要求**：
  - **盟威**：教程第 8 章完整示范了用户角色/用户权限设计（创建经理组、操作员组、用户管理）；其登录与主题采用配置文件式管理（`Config.ini`：`LoginUserList`/`AutoLogin-Enabled`/`ThemeColor`/`LoginInterfaceStyle`，8.3 启动配置可对照）；
  - **Edonsoft 登录**：登录窗体 `USysFrmLogin`（界面 + 口令验证 `VerifyPassword`/`HashPassword`）；
  - **Edonsoft RBAC**：角色管理窗体 `USysFrmRoleManager`（角色→模块权限→用户的 RBAC 模型）+ 权限判定 `GetPermissions`/`LogOff`——见 `Edonsoft Development Framework_x64/edonsoft-docs/13-architecture-source.md` 第 4 节与 `Edonsoft Development Framework_x64/edonsoft-docs/api/07-security.md`，做登录/权限界面与逻辑时直接借鉴其窗体与函数划分。
- **输出物**：角色矩阵 + 用户清单 + 权限配置。

### 8.3 发布、更新与启动配置

- **任务内容**：发布版本、配置启动（AutoExec/启动窗体/启动菜单）、处理更新。
- **任务要求**：
  - **启动配置**：财务系统的启动配置见 `references/开发财务管理系统.pdf` 第 7 章（系统的启动）；
  - **Edonsoft 自动更新**：AutoExec 启动检测版本、替换前端后重启（`USysFrmSettings`/`USysFrmSetup` 配置更新参数，见 `Edonsoft Development Framework_x64/edonsoft-docs/01-quickstart.md` 与 `13-architecture-source.md` 第 8 节第 7/8 行）；
  - **盟威自动升级**：`Update.accde`/`Download.accde` 是平台自带的自动升级机制（见平台教程第 2.4 节文件构成）；
  - **发布新版本**：`vcs_rebuild_database` 从 `.src` 构建全新库（干净构建、部署分发），步骤与成功判别见 `msaccess-vcs-mcp\docs\AGENT_WORKFLOWS.md` 第 6 节，发布检查清单见 `msaccess-vcs-addin\Wiki\PUBLISH.md`。
  - **启动属性/选项的官方写法**：用代码设置启动属性与选项（`set-startup-properties-and-options-in-code`、`set-form-report-and-control-properties-in-visual-basic`、`set-options-from-visual-basic`）见 `..\VBA-Docs\access\Concepts\Settings\`（6 篇）；启动选项 API 单页按名查 `..\VBA-Docs\api\Access.Application.SetOption.md` 与 `Access.Application.GetOption.md`。
- **输出物**：可发布的版本 + 启动配置 + 更新说明。
- **VCS 融入**：**VCS 在运维期持续跟踪**——每次发布前 `scripts/vcs_drive.py` 的 FullExport 子命令导出源码快照，保证可从文本完整重建；发布后用 `scripts/vcs_consistency_check.py` 确认发布版本与文本源对齐；删文件走回收站 `scripts/recycle.py`，不直接删除。

### 8.4 备份与恢复

- **任务内容**：建立运维期备份与恢复机制——文本源即备份、发布快照、后端库备份、恢复演练。
- **任务要求**：
  - **文本源即第一备份**：`.src` 是唯一真相源，可从文本完整重建库——恢复 = `vcs_rebuild_database` 从 `.src` 构建全新库（步骤与成功判别见 `msaccess-vcs-mcp\docs\AGENT_WORKFLOWS.md` 第 6 节）；
  - **发布快照**：每次发布前 `scripts/vcs_drive.py` FullExport 导出源码快照（同 8.3 VCS 融入），发布版本与文本源漂移为 0（`scripts/vcs_consistency_check.py` 核验）；
  - **后端数据备份**：前后端分离后 `Data.accdb`（或链接后端）存共享位置，需单独定期备份；构建过程自动产生的备份库（Build 改名留档的原库）保留最近几份即可；
  - **恢复演练**：按「删除前端 → 从 `.src` 重建 → 重链后端」跑通一次恢复，确认不依赖任何未版本化的手工改动；删文件走回收站 `scripts/recycle.py`，不直接删除。
- **输出物**：备份策略说明 + 恢复演练记录。
- **VCS 融入**：备份与恢复完全基于 VCS 文本源与构建通道——无额外步骤，见 2.1 与 8.3。

### 8.5 收尾核验（工作流的闭环视角）

- **任务内容**：对照开场总览与各阶段输出物自查，确认整条工作流闭环成立。
- **任务要求**：
  - **逐项核对**——每个阶段的输出物都已产出并成为下一阶段输入；
  - `.src` 文本源是唯一真相源且与实时库漂移为 0；构建成功判定四条件成立；发布版本可从文本完整重建；
  - **收尾发现问题时的出口**：按症状查 `docs/troubleshooting.md`；加载项行为疑问（安装/选项/合并/测试的常见问题）查 `msaccess-vcs-addin\Wiki\FAQs.md`。
- **输出物**：交付核验清单。
- **案例的定位说明**：案例不是收尾资源——已按使用场景嵌入各阶段（2.1 闭环演练、2.2/2.3/2.4 建表关系查询活样例、3.1 模块积木、4.1/4.2 窗体控件样例、8.3 业务样例），做哪一步就去看哪一步引用的案例，收尾无需再导航。
