# references/ 导航

本目录是 Access 数据库应用系统开发的「硬规则与方法论」沉淀。它不是索引表，而是一组完整可独立成立的思想单元：遇到哪类问题，就读对应的那一节。

## 与 SKILL.md 的分工

- **SKILL.md 只讲主干**：9 阶段工作流的任务内容/任务要求/输出物/VCS 融入四要素、每阶段该复用哪个积木、每个对象类型在哪个阶段处理——是"怎么做"的连续路线。
- **references/ 展开细节**：SKILL.md 每阶段引用的硬规则、编码字节级细节、脚本参数、命令写法、判别方法，都在这里按问题分文件展开——是"做对"的依据。
- **单点优先、双向导航**：每个内容只在一处展开（单点），其余文件一律只做导航引用——单点可以是 SKILL.md（如「二、对象全景与文本形态」「7.2 自动化测试手段」），也可以是本目录任一文件（`iron-laws.md` 为结论索引、各专文为展开点）；两者结论冲突时以 `references/iron-laws.md` 与 `docs/troubleshooting.md` 为操作口径（实测为准）。

## 按问题读

- 动手前先把编码、类模块、工程定位、弹窗守卫这几条刻进习惯——读 `references/iron-laws.md`。
- 「写进 Access 到底该用什么编码、为什么中文乱码会报 SQL 3075」——读 `references/vba-encoding-rules.md`。
- 「用 Python win32com 怎么驱动 Access、怎么改 VBA 工程、为什么 ActiveVBProject 不是你的工程、Tab 控件怎么建、命令速查」——读 `references/vba-com-automation.md`。
- 「VCS 四个目录实体（源码仓库/完整分发/历史快照/MCP）什么关系、插件与 MCP 怎么安装配置、构建与回写的正确姿势、加载项崩溃怎么判别、NetUI 对话框怎么解锁」——读 `references/vcs-three-channels.md`。
- 「在 `.src` 文本源上改对象（表/查询/窗体/报表/模块）的配对规则、编码换行、安全删除清单」——读 `references/vcs-text-editing.md`；**改哪种对象、处理哪个环节，按对象类型读对应手册（9 份，一对象一份，与本文平级）**：
  - 改**窗体**布局/控件/事件接线/组合框/子窗体宽度——读 `references/forms.md`。
  - 改**报表**布局/节/排序分组/节宽页宽——读 `references/reports.md`。
  - 窗体/报表里的**图片**（增删/替换/OLE 转 Image/ImageData 十六进制格式）——读 `references/images.md`。
  - **条件格式**（解码二进制块 → 删除 → VBA `FormatConditions` 重建）——读 `references/conditional-formatting.md`。
  - 改**查询**（`.bas`+`.sql` 配对、单一 Joins 块、SaveQuerySQL 交互）——读 `references/queries.md`。
  - 建**表/关系/链接**（ID 主键规范、关系 Attributes 值、ODBC 连接串一致）——读 `references/tables-and-relationships.md`。
  - 写 **VBA 模块/类模块/代码后置**（`.cls` 头规则、VBA7、超链接列）——读 `references/vba.md`。
  - 动**工程配置文件**（vcs-options/vcs-index/vbe-project 等 10 个、构建导入顺序）——读 `references/project-config.md`。
  - 改**宏/共享图像/导入导出规格**——读 `references/other-objects.md`。
- 「纠结盟威快速开发平台还是 Edonsoft 框架（能力差异对照/按需取用/与工作流的关系）」——读 `references/framework-comparison.md`。
- 「选底座/部署时纠结 C/S 还是 B/S、Access 文件共享还是上 SQL Server」——读 `references/access-architecture-network.md`。
- 「技能 vendored 了哪些上游开源 / 商业目录、怎么检测它们被手改、怎么合法更新上游」——来源与基线指纹登记见 `manifest.json`，离线漂移检测见 `scripts/check_upstream_drift.py`（`--register` 重算基线、`--no-upstream` 纯离线），门禁拦截见 `tests/test_skill_integrity.py` 第 5 项检查，同步 / 漂移事件流见 `SYNCLOG.md`（原则：vendored 原文件永不手改，更新须整体替换 + 重新 `--register` + 记一条）。

所有「出错了怎么办」按症状去 `docs/troubleshooting.md` 查。

想照一个完整真实案例走一遍（财务管理系统：分析 → 建表 → 查询 → 窗体 → 报表 → VBA → 启动）——看 `references/开发财务管理系统.pdf`（只读参考，章节印证工作流阶段 1（分析设计）、2（表与查询）、3（代码模块）、4（窗体）、5（报表）、8（启动））。

想要一个格式全对的完整样例对照学习——看 `examples/CustomerOrders/`（一个 prompt 生成的订单系统：`CustomerOrders.accdb` 成品 + 全量 `.src`；与上文 9 份手册逐条对照：窗体→`forms/`、报表→`reports/`、查询→`queries/`、VBA→`modules/`、表→`tbldefs/`、关系→`relations/`、配置→顶层 `.json`）。
