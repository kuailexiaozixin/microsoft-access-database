# 两框架差异对照：盟威 Access 快速开发平台 vs Edonsoft 开发框架

> 本文是 `SKILL.md` 阶段 0.2「框架与积木地图」的详尽配套。两框架都不是必选底座，而是**按阶段按需取用的积木来源**——能力、教程、文档在 SKILL.md 各阶段直接引用（见 0.2 的两段映射）。本表回答"两者能力怎么对照、什么场景借鉴哪个实现"。
> 事实来源：盟威侧——平台自带教程 `盟威Access快速开发平台V2.7.0版(64位)/Access2016报销管理系统案例开发教程.pdf`（九章完整案例）；Edonsoft 侧——`Edonsoft Development Framework_x64/edonsoft-docs/framework-overview.md`（总纲）与 `13-architecture-source.md`（架构）。**凡未经真机实测、仅出自教程/文档的表述，表中以『据教程』/『据文档』标注；已实测项直接写明实测结论**——若后续实测与表冲突，以实测为准。

## 一、差异对照

| 维度 | 盟威 Access 快速开发平台（V2.7.0，64 位） | Edonsoft 开发框架（v202607.01） |
|---|---|---|
| **定位** | 企业级快速开发平台：按向导快速出业务系统，少碰底层基础设施 | AI 友好的开发框架：公开函数全部可 `Application.Run` 调用，代码驱动优先 |
| **构成** | `Main.accdb`（客户端主程序）+ `Data.accdb`（后台数据库）+ `Update.accde`/`Download.accde`（自动升级）+ `RDPLib.ucl`（支持库）+ `Config.ini` | 前端宿主库 `EdonSoft Development Framework.accdb` + 核心编译库 `EdonSoftCore.accde` + 数据后端 `Data.accdb` + `appsettings.json`（连接配置）+ `assets\`（主题/前端资源） |
| **前后端分离** | 主程序与数据分离（`Main.accdb` ↔ `Data.accdb`），带自动升级通道 | 宿主库与数据分离（前端库 ↔ `Data.accdb`），核心逻辑编译在 `EdonSoftCore.accde` |
| **开发方式** | **向导驱动**：数据模块向导、导航菜单向导、查阅列表、启动菜单、报表向导（静态/动态/切换面板） | **代码生成器 + 公开 API**：`USysFrmCodeGeneratorPro` 生成业务窗体；60 个公开函数 7 类（登录/通用查询/搜索组合框/后端/升迁/待办/仪表盘/单据选择等） |
| **AI 可驱动性** | 教程导向，窗体/向导为主；自动化需要走 COM 驱动界面与 VBA | **强**：核心库启动时自动引用（AutoExec → `basCoreReference` 修复引用 → 延迟绑定 `Application.Run "Main"`），全部公开函数可 COM `Run` 直调 |
| **自包含自运行** | 随技能目录存放；`反编译打开Main.accdb.vbs` 提供反编译入口（框架库加密，常规开发用向导/成品） | 已验证自包含自运行：前端库/核心库/数据后端同目录即自启动，核心引用自动修复（`Broken=False`），`AutoNumStr` 等 API 实测可调 |
| **升级部署** | 内置 `Update.accde`/`Download.accde` 自动升级机制（据教程） | 无独立升级通道（据文档；`07-upsize.md` 是数据库升迁到 SQL Server 的向导） |
| **权限体系** | 教程第九章「用户权限设计」：权限向导式配置（据教程） | RBAC 角色/权限管理（据文档）：`USysFrmRoleManager`/`USysFrmPermissionSettings` 等 + `VerifyPassword` 登录 API |
| **典型场景** | 不想碰底层、按向导快速出表单+报表的**人工开发**项目（报销系统等典型 MIS） | 需要 **AI 全程代码驱动**（登录/菜单/查询/单据/仪表盘/待办/升迁）的自动化开发项目 |
| **文档/学习路径** | 单本 PDF 九章完整案例（需求→表→查询→窗体→报表→权限），对应工作流各阶段 | `edonsoft-docs/` 14 篇主题文档 + `api/` 60 函数目录（按官方帮助页左侧菜单完整转写） |

## 二、按需取用（什么场景借鉴哪个）

| 场景 | 借鉴盟威 | 借鉴 Edonsoft |
|---|---|---|
| 开发主力 | 人工、向导、界面点击 | AI/脚本代码驱动（`Application.Run`） |
| 升级部署要求 | 需要自动升级通道（Update/Download） | 单机/内网直接复制部署 |
| 界面产出方式 | 向导生成（数据模块/导航/报表向导） | 代码生成器 + 公开 API 组装 |
| 权限复杂度 | 向导式用户权限设计 | RBAC 角色权限（更细粒度） |
| 文档形态 | 单教程九章（线性案例） | 主题文档 14 篇 + API 目录（参考型） |
| 主界面组件 | 查阅列表/启动菜单 | 待办（08）/仪表盘（09） |

## 三、与工作流的关系

两框架都是**积木来源**，不是必选底座：

- 借鉴到的能力落进项目库（表/查询/窗体/模块）后 → **阶段 2.1 起纳入 VCS**（导出 `.src` 作单一真相源）。
- 后续每个增量（表/查询/窗体/模块）都走「改 `.src` 文本 → 构建/回写 → 漂移核验」闭环；框架库本体**只作为参考源**，不直接修改（盟威 `Main.accdb` 加密、Edonsoft 核心在 accde 内）。
- 驱动框架能力：Edonsoft 直接 `Run` 公开函数；盟威走 COM 驱动窗体/向导（见 `references/vba-com-automation.md`）。
