# 示例（examples）

本目录放**实操案例（成型系统或者可复用的组成部分等等）**，分两类：

1. **闭环演示**（`01_`、`02_`）：照着真机跑一遍就能看到技能核心主张成立——**版本控制（VCS 加载项 + msaccess-vcs-mcp）是贯穿开发全生命周期的方法，每一步都走「改文本源 → 构建/回写回库 → 漂移核验」的闭环，全程代码驱动、零手动**。
2. **积木型案例**（`examples/Access DatePicker/`、`examples/Access VBA Modules Collection/`、`examples/Access BOM Management System/`、`examples/AccessAI/`）：从上游仓库带入的成熟可复用组成部分，对应工作流中「控件 / 通用代码库 / 业务结构样例 / AI 能力」四个位置，直接拿去当作你自己系统的构建积木。
3. **完整业务样例**（`examples/CustomerOrders/`）：一个 prompt 生成的订单系统成品（`CustomerOrders.accdb` + 全量 `.src` 文本源），是 9 份对象手册（`references/forms.md` 等）的**活样例**——每种对象的格式规则都能在示例里找到对应实际文件核对（窗体→`forms/fCustomerList.bas`+`.cls`、报表→`reports/rOrderReport.bas`、查询→`queries/qOrders.bas`+`.sql`、VBA→`modules/modNavigation.bas`、表→`tbldefs/tOrders.xml`、关系→`relations/tCustomerstOrders.json`、工程配置→顶层 6 个 `.json`），AI 生成/修改对象前先照此对一遍。

> **发布说明**：本技能发布到公开仓库时，**来自开源仓库的示例目录一律以 `README.md` 占位**——`AccessAI` / `Access BOM Management System` / `Access DatePicker` / `Access VBA Modules Collection` / `CustomerOrders` 五个目录不随仓库分发 `.accdb` / `.mdb` / `.rar` 等重二进制与第三方源码，各自仅保留 `README.md`（写明来源与用途）。本地完整副本用于离线运行与漂移跟踪，清单见 `manifest.json`、同步记录见 `SYNCLOG.md`。

每个闭环演示都是独立的 `run.py`，用 `python run.py` 即可运行（需要本机已装 Access 与技能目录运行位的 VCS 加载项 `Microsoft Access Version Control System\Version Control.accda`）。脚本会自己建临时库、调本技能 `scripts/` 里的工具、跑完清理，不污染你的任何文件。两个闭环演示的定位不同：**01 演示「初始化基线」**（库一建好就纳入版本控制，对应阶段 2.1 的 VCS 融入动作），**02 演示「增量回写」**（改文本源 → 原子重导入 → 运行断言，是阶段 3 改码与阶段 6.2 AI 改库的最小单元）。

| 示例 | 类型 | 对应工作流阶段 | 用途 |
|---|---|---|---|
| `examples/01_blank_db_to_vcs_loop/` | 闭环演示 | 阶段 1/2 | 从空库起步 → 导入一个模块 → 经 VCS 导出 `.src` 作为单一真相源 → 漂移核验证明实时库与 `.src` 对齐 |
| `examples/02_edit_src_reimport/` | 闭环演示 | 阶段 3 | 增量回写：直接在 `.src` 文本上改一处 → 原子重导入回库 → 断言改动生效 |
| `examples/Access DatePicker/` | 积木 | 阶段 4（控件） | 纯 VBA 日期/年月选择器（`Module_DatePicker.bas` / `Module_YearMonthPicker.bas` + `DatePicker.accdb` 演示库 + `使用说明.md`） |
| `examples/Access VBA Modules Collection/` | 积木 | 阶段 2/3/4/5 | 高频模块集合：自动编号（阶段 2 编号字段）、字段校验/ADO/VBE 工具（阶段 3 代码库、阶段 4 控件）、导出 Excel/PPT/HTML（阶段 5 报表输出）；分阶段用法见 `SKILL.md` 对应步骤与 `README.md` + `examples/Access VBA Modules Collection/wiki/` 文档 |
| `examples/Access BOM Management System/` | 积木 | 阶段 2（建模参考） | 多级 BOM 管理样例（`BOM.accdb` / `BOM.mdb` / `treeview.accdb`），父子结构落地的完整参照 |
| `examples/AccessAI/` | 积木 | 阶段 6（AI 接入） | Access 内调用大模型的成熟工具库（Access LLM Toolkit）：`CreateAIForm` 一键建窗体，支持流式输出、对话历史持久化、Access SQL 助手、TXT/CSV/Word/Excel/PDF 文档问答、DPAPI 加密 Key（`AI.accdb` + `JsonConverter.bas` + `Module_Markdown.bas`） |
| `examples/CustomerOrders/` | 完整业务样例 | 全阶段（尤其 2/3/4/5/6.2） | 一个 prompt 生成的订单系统成品（3 表 4 查询 6 窗体 1 报表 + 导航模块 + 全配置）：9 份对象手册的活样例，`CustomerOrders.accdb` 开箱可运行，`.src` 逐对象对照格式规则学习；`prompt.txt` 是生成它的原始 prompt 模板——新库从 0 到 1 时保留句式、替换业务与字段即可 |

想看更完整的分阶段走查与具体命令，回到 `SKILL.md` 与 `references/vba-com-automation.md`。
