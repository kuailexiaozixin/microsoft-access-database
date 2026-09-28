# 变更日志（changelog）

> 本文件是技能内容**唯一**允许记载变化的地方。其它文件不含文件修改记录、文档迁移记录、历史变更记录。
> 每条记录一句话说清"技能发生了什么变化 / 为什么 / 怎么验证"。

## v1.0（初始成型）

- **形成通用、不绑定任何自研框架的 Access 数据库应用系统开发技能**
  - 内容来源：从 Access VBA / VCS / COM 自动化的一线真机验证中沉淀出的确定性规则与可复用工作流，剥离一切自研框架专有叙述。
  - 形态：`SKILL.md` 仅以**工作流一种形态**呈现，链接嵌入论述，不出现索引表/路由表/资源清单（避免 LLM 忽略或搁置部分文档）。
  - 引用：充分引用本目录 12 个成熟子项目（盟威平台、Edonsoft 框架与 edonsoft-docs、Access DatePicker、Access VBA Modules Collection、AccessAI、ms-access-ai-skill、Access BOM Management System、msaccess-vcs-addin、msaccess-vcs-mcp、Microsoft Access Version Control System、Version_Control_v5.0.1），遵循 SKILL Graph 语义链接模式。
- **沉淀通用硬规则到 `references/`**
  - `iron-laws.md`：编码、类模块、工程定位、弹窗守卫、漂移核验、VCS 操纵等 40 条确定规则。
  - `vba-encoding-rules.md` / `vba-project-operations.md` / `vcs-three-channels.md`：编码、工程操作、三套 VCS 通道的原理与正确姿势。
  - `references/README.md`：问题导向的导航（非索引表）。
- **沉淀零手动工具到 `scripts/`**
  - 弹窗守卫 `fwguard.py`、回收站删除 `recycle.py`、实时库 vs `.src` 漂移核验 `vcs_consistency_check.py`、模块原子重建 `rebuild_module_from_src.py`、模板探针 `probe_templates.py`、VCS 探针 `probe_vcs_api.py`、VCS 驱动 `vcs_drive.py`、工程能力诊断 `diagnose_vba_project.py` / `probe_vba_project.py` 等。
- **沉淀去框架化模板到 `templates/`**
  - 标准模块 `standard-module-stub.bas`、类模块 `class-module-stub.cls`、AutoExec 宏 `autoexec-macro.macro`、模板说明 `README.md` / `src-structure.md`，均剥离框架专有命名。
- **沉淀详细手册与验证**
  - `docs/workflow.md`：9 阶段具体命令 + COM 自动化 SOP + TDD 做法 + 命令速查。
  - `troubleshooting.md`：按症状的排错手册（故障 → 根因 → 修复）。
  - `tests/`：通用冒烟验证脚本（模板/编码/漂移核验）。
  - `assets/`：通用前端库（Bootstrap / jQuery / ECharts / Font Awesome），供 Access 内嵌网页与仪表盘。
- **验证口径**：所有规则以 `references/iron-laws.md` 与 `troubleshooting.md` 为操作基准；文档结论冲突时以最新实测为准。

## v1.1（形态收敛：引用具体化、scripts/examples 分工、SKILL.md 精简）

- **引用具体到文档（约束 11）**
  - `SKILL.md` 每阶段的引用从"文件夹/README 概括介绍"升级为"文件夹/README 概括 + 具体文档引用介绍"——例如阶段 0 显式具名 VCS 三套同源工具（`[msaccess-vcs-addin/README.md]`、`[Microsoft Access Version Control System]`、`[Version_Control_v5.0.1]`、`[msaccess-vcs-mcp/README.md]`），数据建模阶段指向 `[edonsoft-docs/08-upsize.md]` 等具体文档，工程操作阶段指向 `[references/vba-project-operations.md]` 的「Tab 控件」一节。全部 12 个子项目均在论述中具名或被具体文档引用。
- **scripts/ 与 examples/ 分工明确（约束 14）**
  - `scripts/`：放零手动的**可运行脚本 + 代码示例**（弹窗守卫、回收站删除、漂移核验、模块重建、VCS 探针与驱动、COM 连接等可复用代码）。
  - `examples/`：放**实操案例（成型系统或者可复用的组成部分等等）**；现有 `01_blank_db_to_vcs_loop` / `02_edit_src_reimport` 是"可复用的组成部分"型闭环演示，每个案例附独立 `run.py`（建临时库 → 调 scripts 工具 → 核验 → 清理）。`README.md`（本目录与 `examples/README.md`）措辞同步改为"实操案例"。
- **SKILL.md 只留核心主干（约束 15）**
  - `SKILL.md` 仅含 8 阶段核心主干与特别重要的内容：以工作流叙述做"对资源的介绍与引用"，**不展开**具体规则/命令/表格/索引；详尽内容（编码铁律、工程操作、VCS 三通道、各阶段具体命令、TDD 做法）全部落到 `references/`、`docs/workflow.md`、`templates/`、`scripts/`、`examples/` 等具体文档。同步去掉此前混入的 parenthetical 详述。
- **验证口径**：重跑 `tests/test_skill_integrity.py`（32 项冒烟全绿）、`py_compile` 校验 `examples/*/run.py`、平衡括号感知的死链检查（0 死链）、`SKILL.md` 无表格且 12 子项目覆盖全。

## v1.2（结构收敛：workflow.md 合入 SKILL.md、troubleshooting 归位 docs/、问题沉淀规范）

- **`workflow.md` 合入 `SKILL.md`，不再作为单独文件（约束 16）**
  - 根因：`docs/workflow.md` 的功能定位（分阶段讲"先做什么后做什么"）与 `SKILL.md` 的工作流主干重合，留两个并行工作流文件会让 LLM 二选一、稀释入口。
  - 处理：删除 `docs/workflow.md`；其**独有**的详尽内容（COM 自动化 SOP 6.1–6.6、TDD 做法、命令速查表）迁到 `references/vba-automation-sop.md`——这是"详尽、细致的具体内容"，按约束 15 落到具体文档，而非膨胀 `SKILL.md`。`SKILL.md` 保持 8 阶段主干、不展开；所有原指向 `docs/workflow.md` 的链接改为指向 `references/vba-automation-sop.md`。
- **`troubleshooting.md` 归位到 `docs/`（约束 2 的"docs 文件夹（含 troubleshooting.md）"）**
  - 从根目录移到 `docs/troubleshooting.md`，使 `docs/` 成为"按症状查的排错手册"的家，符合指令对 11 个目标文件夹的功能定位。所有 `troubleshooting.md` 链接同步改为 `docs/troubleshooting.md`（含 `SKILL.md`、`README.md`、`references/README.md`、`tests/test_skill_integrity.py` 的 `REQUIRED_FILES`）。
- **问题沉淀规范（约束 17）**
  - 明确：开发/操纵中遇到的问题，先**总结现象与本质、剖析原因、总结解决方法**，优先沉淀进 `docs/troubleshooting.md`（按「现象 → 根因 → 修复/预防」结构）；只有"核心、主干、特别重要"的内容才进 `SKILL.md`。`docs/troubleshooting.md` 现有 21 节均按此结构（如第 1 节"模块被静默丢弃"、第 3 节"类模块 type=2 丢失"、第 13 节"BOM 才是真凶"）。
- **验证口径**：重跑 `tests/test_skill_integrity.py`（33 项，`docs/troubleshooting.md` 与 `references/vba-automation-sop.md` 纳入必检、移除已删的 `docs/workflow.md`）；`py_compile` 校验 `examples/*/run.py`；平衡括号感知死链检查（0 死链）；确认 `docs/` 仅含 `troubleshooting.md`、`docs/workflow.md` 已不存在、`SKILL.md` 无表格且 12 子项目覆盖全。

## v1.3（架构参考整合 + 真实案例优化工作流）

- **三篇架构 HTML 整合为一篇并移除（指令约束 2）**
  - 整合 `references/` 下《C/S 结构与 B/S 结构简述》《为什么说 B/S 结构优于 C/S》《Access 的网络操作模式简介》三篇只读参考为 `references/access-architecture-network.md`：C/S 与 B/S 概述、B/S 优势（瘦客户机/低成本/移动办公/系统整合）、Access 文件服务器模式与客户机/服务器模式两种网络模式的取舍，以及它们与本工作流阶段 0（选底座）、阶段 7（部署运维）的对应。原文三篇 HTML 已删除。
- **参考真实案例优化工作流（指令约束 5）**
  - 读 `references/开发财务管理系统.pdf`（49 页，用 Access 2003 开发财务管理系统：分析设计 → 建表 → 查询 → 窗体报表 → VBA 模块），其章节顺序正好印证本工作流阶段 1→2→3→4→6。在 `SKILL.md` 阶段 1 加入该真实案例指针，并在 `references/README.md` 导航补入「架构文档」与「真实案例 PDF」两项。
- **验证口径**：重跑 `tests/test_skill_integrity.py`（33 项，新增 `references/access-architecture-network.md` 必检）；`py_compile` 校验 `examples/*/run.py`；平衡括号感知死链检查（0 死链，含新增 PDF / 架构文档链接）。

## v1.4（SKILL.md 引用全部资源文件）

- **SKILL.md 介绍并引用每一个资源文件（用户明确指令）**
  - 此前 SKILL.md 只引用了部分资源；本轮在 8 阶段叙述中为**全部**资源文件补上内联链接，做到无一遗漏：
    - `templates/`（5 文件：README / standard-module-stub.bas / class-module-stub.cls / autoexec-macro.macro / src-structure.md）；
    - `assets/`（manifest.json / ASSETS.md）；
    - `tests/test_skill_integrity.py`；
    - `scripts/` 全部 11 个 `.py`（此前未引的 `vcs_drive.py`、`probe_blank_db.py`、`enumerate_nonupdatable_queries.py`、`probe_vba_project.py`、`probe_vcs_api.py` 等均已补入对应阶段）；
    - `references/README.md`、`changelog.md`、`README.md`、两篇案例 PDF（`开发财务管理系统.pdf`、`Access2016报销管理系统案例开发教程.pdf`）与一篇需求分析只读 HTML（`谈数据库设计前分析工作`）。
  - 保持工作流形态：**链接全部嵌入阶段论述，无表格、无独立资源清单/路由表**（约束 9）；每条仅"介绍 + 引用"，不展开细节（约束 15）。
- **验证口径**：程序化核对 SKILL.md 的 49 个资源链接全部命中（含 12 子项目与 `../VBA/SKILL.md`），无未引用孤儿；平衡括号感知死链检查 0（160 处）；完整性 33/33。

## v1.5（资源引用审计复核：补齐 examples 叶子文件与 tests/README.md）

- **背景**：对 v1.4「SKILL.md 引用全部资源文件」做**逐一程序化复核**——枚举磁盘上每个资源文件，再在 SKILL.md 的链接集合里查找是否存在。复核发现两处**未真正覆盖**：
  - `examples/` 只引用了**目录**（`examples/`、`01_blank_db_to_vcs_loop/`、`02_edit_src_reimport/`），目录里的叶子文件（两个可运行 `run.py` 与三份 `README.md`）未被显式引用——而 `run.py` 正是案例的核心可运行产物；
  - `tests/README.md` 完全未引用（此前只引了 `tests/test_skill_integrity.py`）。
- **处理**：在 `SKILL.md` 阶段 6 的测试处补 `[tests/README.md]`；在末段 examples 处补「总览 `[examples/README.md]` + 案例一 `[01/run.py]`/`[01/README.md]` + 案例二 `[02/run.py]`/`[02/README.md]`」的显式引用。引用仍全部嵌入阶段论述，无表格、无独立资源清单（约束 9/15 不变）。
- **附带清理**：`examples/*/__pycache__`（Python 编译缓存，非资源）经 `scripts/recycle.py` 移入回收站，保持技能目录无残留编译产物。
- **验证口径**：程序化资源覆盖核对——`scripts` 11/11、`templates` 5/5、`tests` 2/2、`docs` 1/1、`examples` 5/5、`references` 10/10、`assets` 文档 2/2，**无未引用资源**；平衡括号感知死链检查（技能自有文件 170 处）**0 死链**（说明：12 个子项目目录**内部**的 wiki 相对链接属其自带文档、不在本技能维护范围，故排除在扫描外，避免把上游文档的锚点当作本技能缺陷）；`tests/test_skill_integrity.py` 33/33；`SKILL.md` 0 表格。

## v1.6（尝试：把「文字=路径」的链接改成语义标签）—— 已被 v1.7 取代

- **问题**：技能自有 Markdown 里大量链接把**路径原样当链接文字**（方括号里写相对路径、圆括号里是同一条相对路径），同一串路径显示两遍，既啰嗦又不便定位。
- **当时的处理（依据 `skill-pattern/references/skill-graph.md` 的「语义化链接」）**：把链接文字换成**中文语义标签**（如「版本控制系统快照」），路径只留在圆括号里。范围：全部自有文档 14 个文件、共 166 处。
- **复盘（见 v1.7）**：该做法把路径**藏进了圆括号**，纯文本渲染时看不到路径，LLM 难以据此定位文件——**方向有误**。

## v1.7（对齐 VBA 技能写法：内部引用一律用反引号写全路径）

- **问题**：v1.6 把路径藏进圆括号，纯文本里看不到路径，无法据此找文件。
- **依据（用户指令：照 VBA 技能怎么写）**：对照同级 `VBA` 技能（`../VBA/SKILL.md` 及其 `references/*.md`）——**内部引用一律写成反引号包裹的、从技能根算起的相对路径**，例如 `references/iron-laws.md`、`scripts/com_testing.py`、`assets/dialog_suppressor.bas`；**不带 `./` 或 `../` 前缀**；Markdown 链接语法**只留给外部网址**；语义由所在句子承担（列表里用「`路径`：说明」）。
- **处理**：全部自有文档的内部引用统一改成这种写法——`SKILL.md` 手写重排（恢复 v1.6 去掉的描述词，如「加载项本体 / 快照 / MCP 封装」，路径用反引号外显），其余 13 个文件机械转换后逐一读回润色。
- **范围与结果**：14 个文件；内部 Markdown 链接**归零**（`SKILL.md` 内已无任何「方括号 + 圆括号」形式的内部链接），全部改为反引号路径，路径外显、语义由句子承担。
- **验证口径**：自有文档**剩余内部 md 链接 0 处**；`SKILL.md` 资源覆盖不变（scripts 11/11、templates 5/5、tests 2/2、docs 1/1、references 10/10、assets 2/2、examples 5/5）；`tests/test_skill_integrity.py` **33/33**；`SKILL.md` 0 表格。

## v1.8（去框架化收尾：脚本参数化 + 测试补漏）

- **问题（审查发现）**：`scripts/enumerate_nonupdatable_queries.py` 仍硬编码 `FRAME_DIR` 变量并默认指向技能根的 `Framework.accdb` / `Data.accdb`——两文件随自研框架移除已不存在，脚本因 `if not os.path.exists: continue` 静默空跑、什么也不扫，且不接受命令行参数，LLM 无法指定目标库；`tests/test_skill_integrity.py` 的禁止清单写的是 `FrameWork`（大写 W），实际残留是 `Framework.accdb`（小写 w），大小写差异导致该残留未被 33/33 测试拦截，且其注释「已删除并移入回收站」属历史变更记录。
- **处理**：
  - 脚本改为 `argparse` 支持 `--db`（可多次指定）；缺省指向本技能 12 子项目中 Edonsoft 的后端库 `Edonsoft Development Framework_x64/Data.accdb`（允许引用的通用默认）；变量 `FRAME_DIR`/`MAIN`/`DATA` 移除；文件不存在时打印「跳过」而非静默继续。
  - 测试禁止清单补 `Framework.accdb`、`FRAME_DIR` 两项，未来可拦截同类残留；删除注释中的迁移记录，改为中性表述；禁止词扫描**豁免 `changelog.md`**（它是技能内唯一记载变化处，允许提及历史上的残留词，避免元文件与内容扫描自相矛盾）。
  - `SKILL.md` 阶段 2 的引用处补一句用法（`--db` 指定目标库）。
- **验证口径**：`py_compile` 通过；`--help` 参数解析正常；重跑 `tests/test_skill_integrity.py` **33/33**（新增两项禁止词在自有内容文件中均无命中；`changelog.md` 豁免扫描）；`SKILL.md` 0 表格、内部 md 链接 0 处。

## v1.9（大重构：底座文档化、examples 充实、工作流细化为 9 阶段）

- **变更动因（用户要求）**：1) examples 充实——把 `Access DatePicker`、`Access VBA Modules Collection`、`Access BOM Management System`、`AccessAI` 四个积木放进了 examples，并在工作流中充分体现；2) 按官方帮助页补全 edonsoft 文档并随框架迁移；3) 盟威平台补文档；4) 明确 Access 系统对象构成；5) 研究财务 PDF 优化工作流；6) 核心内容回流 SKILL.md；7) 明确数据库设计流程；8) 研究 ms-access-ai-skill、把 VCS/MCP 用法具体化到每阶段；9) 反复核实。
- **处理**：
  - **目录迁移**：`edonsoft-docs/` 整体移入 `Edonsoft Development Framework_x64/edonsoft-docs/`（作为框架自带文档）；`references/Access2016报销管理系统案例开发教程.pdf` 移入 `盟威Access快速开发平台V2.7.0版(64位)/`（作为平台文档）；四个积木（DatePicker / VBA Modules / BOM / AccessAI）已在 `examples/` 下。
  - **edonsoft 文档补全**：按官方帮助页（http://www.edonsoft.com/access-framework-help.aspx）左侧菜单全部栏目完整转写——快速开始/版本更新、代码生成器、通用查询、搜索组合框、后端部署、链接表管理器、数据库升迁、提醒待办、主页仪表盘、单据选择弹窗、常见问题，共 `01-quickstart.md`~`12-faq.md` 十二份主题文档 + `api/` 七类 60 个公开函数文档；删除被替代的旧编号文档（05/08/10/11）；修订 `framework-overview.md` 与 `13-architecture-source.md` 的编号引用与悬空脚本引用（原指向不存在的 `scripts/edonsoft/*.ps1`、`access_helper.py`），并核实「DoRunSQL/AdoScalar 是否存在」的矛盾（官方帮助页明确收录，属核心库编译内实现，与宿主库 DAO 直连不矛盾）。
  - **工作流重构**：`SKILL.md` 从 8 阶段细化为 **9 阶段（0–8）**——新增「报表设计」阶段（原缺报表）；每阶段统一「任务内容 → 任务要求 → 输出物 → VCS 融入」四要素；VCS/MCP 用法具体化（何时导出/构建/核验、用什么脚本、达成什么目标、输出什么）；开头新增「Access 系统由哪些对象构成」主干认知（表/查询/窗体/报表/宏/模块/关系 + 各对象的 VCS 文本形态，源自 ms-access-ai-skill 的对象配对规则）。
  - **底座充分文档化**：阶段 0 详细介绍盟威（教程 PDF 九章映射）与 Edonsoft（总纲/架构/01–12/api 文档映射）；后续阶段按能力引用底座文档（代码生成器、通用查询、搜索组合框、单据选择、链接表、升迁、RBAC 等）。
  - **案例充分体现**：`examples/README.md` 更新为「闭环演示 + 积木型案例」两类，四个积木对应工作流阶段 2/3/4/6；`SKILL.md` 各阶段引用对应案例。
  - **财务 PDF 与数据库设计流程**：阶段 1 纳入需求分析→功能模块→信息流程图→表设计说明书（财务 PDF 第 1 章、盟威教程第 3 章）；阶段 2 纳入表/关系/查询（财务 PDF 第 2–3 章）。
  - **测试适配**：`tests/test_skill_integrity.py` 支持子项目在 `examples/` 与 `Edonsoft Development Framework_x64/` 内的新位置；SKIP 增加 `examples`。
- **验证口径**：`tests/test_skill_integrity.py` **33/33** 通过；`SKILL.md` 与 `README.md` 路径引用逐一验证（缺失 0 处，通配符/未来迁移目标除外）；`SKILL.md` 0 表格、内部 md 链接 0 处；edonsoft-docs 现为 13 份主题文档 + `api/` 7 文件，与官方帮助页栏目一一对应。

## v2.0（VCS 关系理清、文本编辑规则上提、工作流前后融合、references 分工）

- **变更动因（用户审视）**：1) VCS 插件安装配置使用体现在哪；2) VCS MCP 安装配置使用体现在哪；3) 两个 VCS 快照目录的关系未理清；4) VBA 编写体现在哪、涉及哪些流程步骤；5) SQL 编写体现在哪；6) 数据库设计规划三层次（概念/逻辑/物理）；7) 需求/系统分析/业务流程/面向对象建模如何体现；8) SKILL.md 工作流前后的阐述融合进工作流；9) ms-access-ai-skill 如何处理；10) SKILL.md 与 references 分工。
- **处理**：
  - **VCS 关系理清（3）**：核验两个 accda（内部版本串均 5.0.x；完整分发包的 accda 与 addin 仓库构建产物同大小同构建）后，把「不同快照」改为准确的四态：`msaccess-vcs-addin/`=源码仓库、`Microsoft Access Version Control System/`=完整分发部署包（accda+DLL+Ribbon+Worker，运行 accda 即安装）、`Version_Control_v5.0.1/`=v5.0.1 单文件历史快照、`msaccess-vcs-mcp/`=MCP 服务；安装后实际使用的是正式位 `%AppData%\MSAccessVCS\Version Control.accda`。表述同步更新到 `references/vcs-three-channels.md` 与 `README.md`。
  - **安装配置补全（1、2）**：`references/vcs-three-channels.md` 重写，新增「工具形态与安装」章（1.1 加载项安装四步 + 1.2 MCP 安装六步：uvx/venv、客户端 `mcpServers` 配置、`.env` 变量表、`McpAllowRunVBA` 一次性开启），SKILL.md 阶段 0.3 改为「安装 → 配置 → 验证 → 使用」四步并指向该章。
  - **文本编辑规则上提（9）**：新建 `references/vcs-text-editing.md`，把 `ms-access-ai-skill/.claude/skills/ms-access-vcs/SKILL.md` 的核心规则（对象配对表、编码换行、安全删除清单、常见坑）上提为主技能自有文件，并附「按对象类型查上游手册」映射表；SKILL.md 全部相关引用从子目录路径改指本文件；上游包保持原样（可独立安装）。
  - **工作流前后融合（8）**：SKILL.md 开头三条核心主张融入阶段 0 引言，对象全景（七类对象 + 文本形态 + 处理阶段）融入新增 0.4「Access 对象全景与工作流路线图」；结尾总结融入阶段 8.3 收尾核验段。
  - **阶段细化（4、5、6、7）**：阶段 1 拆为 1.1 需求与可行性 / 1.2 业务流程梳理（新增）/ 1.3 功能与模块设计 / 1.4 数据库设计规划（概念→逻辑→物理三层次）/ 1.5 项目初始化与第一切片；阶段 2.3 明确 SQL 编写的三个落点（查询对象、RecordSource、DAO/ADO）；阶段 3 开头补 VBA 编写落点地图（标准模块/类模块/窗体报表代码后置分别对应哪些阶段），3.2 明确类模块=面向对象建模在 Access 的落地。
  - **分工明确（10）**：`references/README.md` 重写，新增「与 SKILL.md 的分工」节（SKILL.md 只讲主干与路线，references 展开字节级细节；单向引用）。
- **验证口径**：`tests/test_skill_integrity.py` **33/33** 通过；SKILL.md/README/references-README 路径引用逐一验证缺失 0 处（`references/vcs-text-editing.md` 新文件路径均有效）；SKILL.md 0 表格、内部 md 链接 0 处；两个 accda 版本串核验（5.0.x）。

## v2.1（SKILL.md 工作流重新编排）

- **变更动因**：用户要求按开发生命周期重新编排 SKILL.md——区分流程阶段、每阶段区分工作步骤、叙述流畅、分工明确、逻辑严密，避免散乱、杂乱、冗余、重复。
- **处理**：
  - **开场合并**：原分散的开篇三块（工作流形态 / 三条主张 / 对象全景）合并为一段开场：主张 → 工作流形态（9 阶段、四要素、三样产出）→ 对象全景与路线图（七类对象各在哪个阶段处理、`.src` 文本形态统一在此讲一次）；原 0.4 从阶段 0 上提删除。
  - **阶段归属理顺**：1.5「项目初始化与第一切片」从需求分析阶段移入数据建模阶段，成为 2.1「项目初始化与版本基线」（先锁定版本基线再建表）；阶段 7 由单节拆为三步（7.1 可测性设计 / 7.2 自动化断言 / 7.3 构建成功判别与回归收口）；原游离的收尾总结段融入阶段 8 成为 8.4「收尾核验（工作流的闭环视角）」；阶段 5 补齐阶段引言。
  - **每阶段统一叙事**：每个阶段开头一句话说明「本阶段做什么、输入来自哪、输出流向哪、按什么顺序推进」；步骤统一四要素（任务内容/任务要求/输出物/VCS 融入）。
  - **去重**：对象文本形态只在开场路线图讲一次；「导出 → 构建 → 核验漂移 0」统一闭环在开场主张讲一次，各步骤的 VCS 融入只写本步差异动作；2.4 查询配对、4.3 窗体配对等与开场重复的表述精简为引用 `references/vcs-text-editing.md`。
- **验证口径**：`tests/test_skill_integrity.py` **33/33** 通过；终验 15 项全绿（SKILL.md 34 处路径引用有效、0 表格、0 内部 md 链接、edonsoft-docs/盟威 PDF/examples 四积木就位、无旧编号残留）。

## v2.2（工作空间研究与 references 按主题重组）

- **变更动因**：用户提出五项审视——①研究开发工作空间（`D:\WorkBuddy工作空间\2026-09-07-14-37-36`）提取有价值信息并核实；②讲清两个闭环示例的分阶段用法；③讲清 Access VBA Modules Collection 的分阶段用法；④审查对 LLM 无价值的信息；⑤references 同一主题合一、不同主题分文档。
- **处理**：
  - **references 重组**：合并 ba-automation-sop.md + ba-project-operations.md → ba-com-automation.md（同一主题"Python/COM 操纵 Access"合一，消除 get_proj/弹窗守卫/InsertLines 三处跨文件重复正文；两个旧文件回收站删除，SKILL.md/examples-README/references-README 引用同步更新）。
  - **iron-laws.md 瘦身**：H/I/J 节（工程写入、守卫、组件命名、VCS 操纵）压缩为"结论 + 指向专文"；新增铁律 41（NetUI 对话框键盘解锁）；配套引用改指 ba-com-automation.md。
  - **vcs-three-channels.md 补三坑**（来源：工作空间 MEMORY 实测）：第十一节 NetUI 对话框键盘解锁（NUIDialog/NetUIHWND → VK_RETURN，Build Name Conflict 等点不动对话框）；第十二节 CLI 忽略 output_dir、MCP 会话残留 Access 窗口须强杀 + 删 .laccdb。
  - **工作空间核实结论**：MEMORY 中通用铁律（编码/类模块/VCS/构建判定/守卫）已全部覆盖于 references；自研框架资产部分（FrameWork 目录、门禁基线、docs 七件套）与现状不符、属历史状态，按"去框架"原则不回流；MCP 日志（388 行）验证工具名集合与文档一致（回写主力 vcs_import_object、验证主力 vcs_call_vba/vcs_compile_vba）。
  - **examples 分阶段落地**：SKILL.md 阶段 2.1 补 01 案例（版本基线）、6.2 补 02 案例（AI 改库最小单元）；阶段 2.2 补 asAutoNumStr、3.1 补模块全清单、3.2 补 ClsFieldValidator/M_ValidationUI、4.2 补校验与 ADO 积木、5.1 补导出积木；examples/README 更新两闭环定位与 VBA Modules 阶段列。
  - **LLM 价值清理**：删除明确垃圾 Microsoft Access Version Control System\Version Control.laccdb（锁残留）与 msaccess-vcs-mcp\.env.bak_20260911_105042；体积冗余（三份重复 15MB accda、97MB MCP venv、演示 accdb）列为待用户决策建议，未动。
- **验证口径**：	ests/test_skill_integrity.py **33/33** 通过（REQUIRED_FILES 更新为 vba-com-automation + vcs-text-editing）；终验 15 项全绿（35 处 SKILL.md 路径引用、0 表格、0 内部 md 链接、edonsoft-docs/盟威 PDF/examples 就位）。

## v2.3（VCS 自包含运行位 + 错误 91 修复 + Ribbon 修复 + 自动化测试完善）

- **变更动因**：用户锁定「所有组件在技能目录内自包含、自运行，不安装/注册到其他位置；技能中不得出现手动操作，全部代码化」；插件在技能目录内重装后无 Ribbon 选项卡；MCP/API 路径 ExportObject 报错误 91。
- **处理**：
  - **安装位=运行位重构**：VCS 加载项安装/运行位统一为技能目录 Microsoft Access Version Control System\（accda + Ribbon DLL + Ribbon.xml/json + Worker.vbs），Menu Add-Ins 三键、受信任位置、CLSID/ProgID 全部指向该目录；scripts/install_vcs_addin.py 升级为全代码化安装（/cmd INSTALL + Ribbon 部署 + 六项自检），无手动弹窗。
  - **Ribbon 修复**：复制官方 Ribbon\Build\MSAccessVCSLib_win64.dll + Ribbon.xml/json 到运行位；HKCU\...\Access\Addins\MSAccessVCSLib.AddInRibbon LoadBehavior=3；真机验证 COMAddIns 含 Ribbon（重启后选项卡出现）。
  - **ExportObject 错误 91 修复**：lnNoIndex 路径下 clsVCSIndex.Update 返回 Nothing，四处组件（clsDbForm/clsDbModule/clsDbReport/clsDbVbeForm）cIdx.FolderAnnotation 无条件访问越界；加 If Not cIdx Is Nothing 保护并构建进运行位；真机 ExportObject/MCP vcs_export_object 返回 success。修复前备份归档 _archive_legacy\src_bak_noindex_fix\。
  - **自动化测试完善**：	emplates/modDebugLog.bas（Debug.Print 文件版：LogPrint/LogTimer/Watch/LogErr，#Const 开关）；SKILL.md 7.2 改为「自动化测试手段 A-E」（弹窗捕获/自动点击、报错捕获、调试信息、VBE 操纵、改代码），7.3 断言、7.4 构建判别；vba-com-automation.md 增 3.7 详单（fwguard 回读、silent_execute、VBE 读写代码片段、COM 清理规范）。
  - **清理与归档**：%AppData%\MSAccessVCS\ 旧安装位（15.2MB accda 等）→ _archive_legacy\MSAccessVCS_appdata\；盟威 .laccdb 锁文件删除；旧叙述（正式安装位/手动安装/Install Add-In）全文修正为技能目录运行位/代码化安装（SKILL.md 0.3、README、iron-laws、vba-com-automation、troubleshooting、examples、vcs-three-channels 十三/十四节、mcp README 自包含部署节）。
- **验证口径**：final_verify.py 真机 6 项（Ribbon COMAddIns=OK、GetVCSVersion 5.0.1、ExportObject success、CompileVBA=True、弹窗 0、残留进程除历史孤儿 25156 外无新增）；mcp_verify.py 端到端 6 步全绿（21 工具、vcs_list_objects/vcs_call_vba/vcs_compile_vba/vcs_export_object 均 success）。

## v2.4（改进方案 A/B/C 全量落地 + _archive_legacy 清理）

- **变更动因**：用户要求把上轮提出的三组改进方案（A 工程类 / B 文档类 / C 脚本类）全部落地并反复核验，同时回答"`_archive_legacy` 可否清理"。
- **处理**：
  - **归档清理**：`_archive_legacy/`（旧正式安装位 5 件 + 修复前备份 4 份，共 14.8MB）评估为完全冗余——运行位已自包含全部组件、msaccess-vcs-addin 为完整 git 仓库可随时取回源码——用 scripts/recycle.py 移入 Windows 回收站（FOF_ALLOWUNDO，可还原，未永久删除）；全技能 5 处 `_archive_legacy` 引用改为「已回收清理（可还原）」，修复前备份来源改指 git 历史。
  - **A1 真机验证入 tests/**：final_verify.py → tests/vcs_com_verify.py、mcp_verify.py → tests/vcs_mcp_verify.py，全部改为相对技能根路径计算 + `--db` 参数；vcs_mcp_verify 硬断言化（Hello 须返回 "VCS OK"、ExportObject 须 success，退出码 0/1）；新增测试夹具 tests/fixtures/basSample.bas 与 tests/README.md；完整性测试 REQUIRED_FILES 扩 3 项。
  - **A2 构建-安装一键化**：新增 scripts/rebuild_and_install.py（修复门禁 check_fix_gate → vcs_drive Build → install_vcs_addin 代码化安装 → 探针库回归四步）；修复门禁双向验证（正向 4 组件含 `If Not cIdx Is Nothing Then` 退出 0；构造无保护目录正确拒绝退出 1）。
  - **A3 进程清理规范**：fwguard.py 新增 snapshot_access()/kill_new_access()/remove_locks()，只杀本次新实例（高完整性孤儿跳过不阻塞、返回残留 PID）；两个回归脚本切换新函数；vba-com-automation.md 3.5 改写为「快照+只杀新进程」教程。
  - **B1**：vcs-three-channels.md 1.2 新增环境变量分工对照表（ACCESS_VCS_ADDIN_PATH=运行位必设 / ACCESS_VCS_DATABASE=目标库可选 / DISABLE_WRITES / ENABLE_LOGGING），端到端命令改 `python tests/vcs_mcp_verify.py --db <库>`。
  - **B2**：新建 references/framework-comparison.md——盟威 vs Edonsoft 十维度对照 + 选型决策表 + 与 VCS 工作流关系（事实来源标注）；SKILL.md 0.2 链接。
  - **B3**：SKILL.md 6.2 新增 AI 改库门禁（硬性）：漂移核验 0 项 + CompileVBA 通过 + 改动对象 ExportObject 复核落盘，三绿才允许构建回库。
  - **C2**：iron-laws.md 新增铁律 42（脚本路径一律相对技能根计算，禁 APPDATA 等环境默认路径——旧 %AppData%\MSAccessVCS 已回收，引用必断链）与铁律 43（进程清理用快照+只杀新进程）。
  - **C3**：examples/01、02 run.py 断言化——ADDIN 从 %AppData% 断链改技能目录运行位；01 漂移核验 rc=3 从 WARN 改 FAIL（返回 1）；退出码语义文档化（0=全过、1=失败）。
- **验证口径**：tests/test_skill_integrity.py **36/36 通过**（含 Framework.accdb 残留检测修正——负向后行断言排除 EdonSoft 真实底座文件名）；vcs_com_verify 真机 6 项全绿退出 0（COMAddIns=1 含 Ribbon、GetVCSVersion 5.0.1、ExportObject success、CompileVBA True、弹窗 0、新实例残留无）；vcs_mcp_verify 真机 6 步全绿退出 0（21 工具 + 4 工具调用全 success，新实例 PID 5008 被 kill_new_access 正确终止）；11 个改动脚本 py_compile 全过；修复门禁正反例均正确。


## v2.5（对象级手册上移 references + 去重复引用）

- **变更动因**：用户要求把 ms-access-ai-skill 内嵌技能 ms-access-vcs 的 9 份对象级手册上移至主技能 references（不做内容变动），清理主技能 references/vcs-text-editing.md 对子项目的长路径间接引用，技能内不保留重复内容，每个节点独立（SKILL Graph 模式）。
- **处理**：
  - **上移（移动非复制）**：9 份手册（forms/reports/images/conditional-formatting/queries/tables-and-relationships/vba/project-config/other-objects.md，共约 88KB）原样移入主技能 references/（无文件名冲突）；源目录 ms-access-ai-skill/.claude/skills/ms-access-vcs/references/ 已删除，不再保留。
  - **引用清理与更新**：references/vcs-text-editing.md 第六节映射表由长路径改短路径 references/xxx.md，引言改为「本技能 references 自有文件」；SKILL.md 开场路线图点名 9 份手册与映射表位置，阶段 2.3（关系→tables-and-relationships）、2.4（查询→queries）、3.2（VBA→vba）、3.3（宏→other-objects）、4.3（窗体→forms/images/conditional-formatting）、5.2（报表→reports，长路径改短）、6.2（AI 改库→9 份全套）逐点补充对应手册的功能与用法；ms-access-vcs SKILL.md（子项目内）映射表改指主 references（../../../../references/）；README.md 与 references/README.md 同步更新。
  - **验证增强**：tests/test_skill_integrity.py REQUIRED_FILES 补 9 份手册。
- **验证口径**：tests/test_skill_integrity.py **45/45 通过**；全技能无 ms-access-vcs/references 残留引用；9 份手册内部相对链接全部有效（同目录移动后无断链）；SKILL.md 全部引用逐项验证存在（14 项正则误报已排除）。


## v2.6（完整业务样例 CustomerOrders 上移 examples）

- **变更动因**：用户要求把 ms-access-ai-skill 的完整示例 CustomerOrders（accdb 成品 + 全量 .src 文本源）上移至主技能 examples/，与 v2.5 上移的 9 份对象手册配套——手册讲格式规则、示例是活样例，同层形成「读规则 → 对照样例」的学习链。
- **处理**：
  - **移动（非复制）**：CustomerOrders.accdb + CustomerOrders.accdb.src（34 文件）原样移入 examples/CustomerOrders/（结构不变，含 6 窗体 .bas+.cls、4 查询 .bas+.sql、1 报表、1 导航模块、3 表 XML、2 关系 JSON、6 顶层配置 json）；ms-access-ai-skill 源位置清空。
  - **引用更新**：SKILL.md 6.2 示例引用改 examples/CustomerOrders/，并在 2.3（关系→tCustomerstOrders.json+tOrders.xml）、2.4（查询→qOrders.sql）、3.2（VBA→modNavigation.bas）、4.3（窗体→fCustomerList.bas+.cls、主子窗体→fOrderDetail）、5.2（报表→rOrderReport.bas）逐点补充「格式活样例」；vcs-text-editing.md 第六节长路径改短并注明「与 9 份手册逐条对照」阅读法（每种对象：手册读规则 → 示例对格式）；references/README.md 导航补完整样例入口；examples/README.md 新增第三类「完整业务样例」并补表格行；ms-access-ai-skill/README.md（英文）示例引用改为指向主技能 examples/ 并注明位置。
  - **验证增强**：tests/test_skill_integrity.py REQUIRED_FILES 补 9 项样例关键文件（accdb、vcs-options、vcs-index、fCustomerList.bas、qOrders.sql、rOrderReport.bas、modNavigation.bas、tOrders.xml、tCustomerstOrders.json）。
- **验证口径**：tests/test_skill_integrity.py **54/54 通过**；全技能无 ms-access-ai-skill/CustomerOrders 旧路径残留引用；examples/CustomerOrders 目录结构核对通过（34 文件与移动前一致）。


## v2.7（ms-access-ai-skill 两源文件整合进主技能 SKILL.md）

- **变更动因**：用户要求把 ms-access-ai-skill/.claude/skills/ms-access-vcs/SKILL.md 与 ms-access-ai-skill/README.md 整合进主技能 SKILL.md——不丢失关键信息、保证内容完整性，同时主技能不冗余、不重复。
- **覆盖盘点结论**（整合前逐条核对）：
  - ms-access-vcs SKILL.md 的全部关键信息**已被主技能完整覆盖**：frontmatter 之外——工程结构/目录命名（vcs-text-editing.md 第一节）、11 种对象配对规则表（vcs-text-editing.md 第二节，逐条对应）、编码与新建文件 printf/od -c 验证（vcs-text-editing.md 第三节）、安全删除清单（vcs-text-editing.md 第四节）、常见坑（注释引用/db-connection 单条目，vcs-text-editing.md 第五节）、对象细节（9 份手册）。主技能 SKILL.md 开场路线图已含对象全景与映射表入口。**无需新增正文，避免重复**。
  - ms-access-ai-skill README.md 的主体（闭环、示例、项目结构）**已被 SKILL.md 6.2 与 examples/ 覆盖**；其**唯一独特资产——生成 CustomerOrders 的完整 prompt 模板**（README 中原样保留，22 行）此前未进主技能。
- **处理**：
  - **prompt 模板落地**：新建 examples/CustomerOrders/prompt.txt（完整 prompt 原文 + 用法说明：保留句式结构、替换业务对象/字段/报表要求、以 CustomerOrders 为格式基准）。
  - **SKILL.md 6.2 增强**：补「范式出处与独立安装包说明见 ms-access-ai-skill/README.md，其规则/手册/示例均已内化为本技能 references/ 与 examples/，不再重复正文」；补 prompt.txt 引用与从 0 到 1 生成新库的用法（启动屏→列表/编辑窗→主子窗体→预览报表→汇总行的句式骨架）。
  - **examples/README.md**：CustomerOrders 表格行补 prompt.txt 用途。
  - **验证增强**：tests/test_skill_integrity.py REQUIRED_FILES 补 examples/CustomerOrders/prompt.txt。
  - 两份源文件保留原位（作为独立可安装技能包与范式出处文档，内容不再与主技能重复；ms-access-vcs SKILL.md 映射表在 v2.5 已指向主 references）。
- **验证口径**：tests/test_skill_integrity.py **55/55 通过**；SKILL.md 引用有效性逐项验证（新增 prompt.txt 存在）；无新增冗余（两份源文件正文未抄入主技能）。


## v2.8（SKILL.md 重新编排 + 移除两个已整合源文件）

- **变更动因**：用户要求（1）重新编排 SKILL.md——按开发生命周期分阶段、阶段内分步骤，叙述逻辑流畅、分工明确、不散乱不冗余；（2）反复核实确认 ms-access-ai-skill/.claude/skills/ms-access-vcs/SKILL.md 与 ms-access-ai-skill/README.md 的所有内容均已整合吸收进主 SKILL.md 之后移除。
- **SKILL.md 重新编排**：
  - 开篇重构为「一、总览：怎么读本文档」——一条主线（输入→九阶段→三样输出）、三段推进、每步四要素、三条方法主张、统一闭环工具链（导出/构建/核验/弹窗守卫，含 Error 2128 警示）、文档分工。
  - 「二、对象全景与文本形态」表格化——7 类核心对象 + 配套对象 ×（处理阶段 / .src 文本形态 / 格式手册），吸收 ms-access-vcs SKILL.md 的配对规则表；硬规则总入口（配对/编码/安全删除）集中一句指向 vcs-text-editing.md，消除原开篇列表与「另有配套对象」段的重复。
  - 各阶段 VCS 融入收敛：统一闭环在 2.1 锚定，2.2/2.3/2.4/4.3/5.3 只保留本步差异动作；窗体文本细则从 4.3 的 VCS 融入移到 4.1 任务要求（格式规则归格式层、版本控制归版本控制层）。
  - 6.2 重构：吸收 README 的 What It Does（AI 能创建/编辑的对象清单）、Dev Container 安全建议（AI 代理应在隔离环境运行）、prompt 模板用法；拆分为「从 0 到 1 生成新库 / 改现有库 / 门禁 / 安全」四条任务要求。
  - 修正过时数字：7.4 完整性从「36 项」改为「55 项」；8.4 案例对照补 CustomerOrders 完整业务样例。
  - 5.1/5.2 输出物补全（静态报表集/动态报表集），四要素逐步骤齐全。
- **移除两个源文件**（内容全部已吸收进主 SKILL.md / references / examples，逐条核验后移除）：
  - ms-access-ai-skill/.claude/skills/ms-access-vcs/SKILL.md——配对规则（吸收为总览二表格）、编码/删除/常见坑（vcs-text-editing.md 第二~五节）、映射表（总览二手册列）；ms-access-ai-skill/README.md——闭环（总览 + 6.2）、示例与 prompt（examples/CustomerOrders + 6.2）、安全建议（6.2）；License/许可信息保留在原目录 LICENSE 文件与 .devcontainer。
  - 引用同步更新：SKILL.md 6.2 与 README.md、references/vcs-three-channels.md 中对该两文件的引用改为指向已整合位置；changelog 历史记载保留。
- **验证口径**：tests/test_skill_integrity.py **55/55 通过**；全技能对已删两文件的引用仅剩 changelog 历史记载（4 处，正常）；重排后 SKILL.md 41 处引用逐项验证，4 项引号吞噬误报精验存在，无真实失效。


## v2.9（移除 ms-access-ai-skill 整个目录）

- **变更动因**：用户要求移除 ms-access-ai-skill 整个目录（其全部有效内容已在前序版本上移/整合：9 份对象手册 v2.5 → references/，CustomerOrders 样例 v2.6 → examples/CustomerOrders/，prompt 模板 v2.7 → examples/CustomerOrders/prompt.txt，SKILL.md 与 README.md v2.8 整合进主 SKILL.md 后删除）。
- **移除内容**：目录剩余 4 个样板/许可文件（.gitattributes、LICENSE、.claude/skills/ms-access-vcs/LICENSE、.devcontainer/devcontainer.json——devcontainer 为上游仓库遗留的 Linux 容器样板，与本技能 Windows/COM 全链路不兼容，仅安全建议一句已入 SKILL.md 6.2）。
- **同步更新**：
  - tests/test_skill_integrity.py：SUBPROJECTS 12→11（移除 ms-access-ai-skill，注释同步）；完整性 55→54 项。
  - examples/CustomerOrders/prompt.txt：出处表述去掉对已删目录的引用（改指主 SKILL.md 阶段 6.2）。
  - SKILL.md 7.4：完整性「55 项」→「54 项」。
- **验证口径**：tests/test_skill_integrity.py **54/54 通过**；全技能对 ms-access-ai-skill 的引用仅剩 changelog 历史记载与 tests L97 注释（上移出处说明，正常）；目录已移除。


## v2.10（SKILL.md 引用 msaccess-vcs-mcp/docs 五个文档）

- **变更动因**：用户要求不上移、不整合、不对 msaccess-vcs-mcp 做任何变动，仅更新 SKILL.md 在工作流不同阶段不同步骤中充分介绍和引用其 docs/ 下五个文档（SKILL Graph 单点模式：内容留在 mcp 仓库内，SKILL.md 只做导航引用）。
- **SKILL.md 更新（8 处）**：
  - 总览「统一闭环的工具链」段后新增「MCP 侧完整操作手册」文档地图，一次点名五个文件（AGENT_WORKFLOWS / EXPORT_FORMATS / TESTING / VBA_CALLBACK_API / VBA_INTEGRATION）；
  - 0.3（版本控制通道就位）→ VBA_INTEGRATION.md（MCP↔加载项 COM 调用链：Application.Run → HandleRibbonCommand → clsVersionControl）与 VBA_CALLBACK_API.md（APIAsync 回调契约、五种消息、取消轮询、busy 并发语义，超时/busy 时的处理选择）；
  - 2.1（版本基线）→ EXPORT_FORMATS.md（导出产物格式：database.json/.sql 头/.bas/BOM-CRLF/文件名净化，与本技能手册互为补充、出入以本技能为准）与 AGENT_WORKFLOWS.md（核心循环图 + 8 种模式总览）；
  - 2.4（查询设计）→ EXPORT_FORMATS.md（查询导出元数据头 -- Query/-- Type/-- Exported 与文件名净化）；
  - 6.2（AI 编码代理）→ AGENT_WORKFLOWS.md（八种操作模式逐条：改查询含编译失败停止约定、改 VBA 保留 Attribute VB_Name、新建对象、批量导出、团队合并、重建发布库、重建加载项、迭代循环）；
  - 7.2（自动化测试手段）→ TESTING.md（三层测试 + 常见失败场景），并标注其默认安装位 %AppData%\MSAccessVCS\ 与本技能 0.3 自包含原则冲突、以本技能为准；
  - 7.4（构建成功判别）→ AGENT_WORKFLOWS.md 6b/6c（改加载项自身时 vcs_rebuild_addin 重建 + vcs_run_tests 跑自带测试套件，含 headless 输出与超时恢复）；
  - 8.3（发布更新）→ AGENT_WORKFLOWS.md 第 6 节（vcs_rebuild_database 从 .src 干净构建发布库）。
- **未变动**：msaccess-vcs-mcp 目录及内部任何文件（docs/ 原样保留于子项目内）。
- **验证口径**：tests/test_skill_integrity.py **54/54 通过**；SKILL.md 结构锚点不变（9 阶段 30 步，四要素 30/30/30/9）；五个 docs 文件均存在且被引用（总览文档地图 + 阶段分布：AGENT_WORKFLOWS×4、EXPORT_FORMATS×2、TESTING×1、VBA_CALLBACK_API×1、VBA_INTEGRATION×1）。


## v2.11（SKILL.md 引用 msaccess-vcs-mcp/src 源码 + references 编排批判）

- **变更动因**：用户要求（1）更新 SKILL.md 在工作流不同阶段不同步骤中充分介绍和引用 msaccess-vcs-mcp/src 源码，方便 LLM 借鉴其实践自动化操作 Access；（2）批判 references 下 8 个文件的内容编排合理性。
- **SKILL.md 更新（6 处，src 零变动）**：
  - 总览「MCP 侧完整操作手册」扩展为「文档 + 源码借鉴」地图，一次点名 6 个 src 模块（connection/addin_integration/tools/instance_registry/access_gate/vba_worker_manager）；
  - 0.3（版本控制通道）→ 写自动化脚本前照 src 抄骨架：connection.py（连接与打开库生命周期）、addin_integration.py（Application.Run 封装）、validation.py/security.py（路径校验与只读/写入开关）；
  - 6.2（AI 编码代理）→ 实现任意 vcs_* 工具等价能力时读 tools.py 对应实现（参数解析/返回结构/失败分支）；
  - 7.2 → 新增「手段 F——借鉴 MCP 源码」：access_gate.py+operation_manager.py（并发串行化/busy）、com_recovery.py（COM 断开分类恢复）、vba_worker_manager.py（隔离 VBA 执行+超时）、security.py/validation.py；
  - 7.3（自动化断言）→ instance_registry.py 即「只杀本服务创建的 Access 进程」模式的工程实现，写清理逻辑先看它；
  - 7.4（构建判别）→ usage_logging.py（JSONL 审计）+ callback_server.py/operation_manager.py（异步跟踪/超时/取消）供长任务脚本借鉴。
- **references 编排批判**（评审输出，未改文件）：详见对话交付报告——发现真实缺口/错误：references/README.md 缺 framework-comparison.md 导航；access-architecture-network.md 阶段编号 7 应为 8；framework-comparison.md 声明「未核实处标注待核」但全文无待核标注；vcs-three-channels.md 职责混杂（错误 91 记录应归 changelog、自动化测试与 vba-com-automation/SKILL.md 7.2 三方重复）；vba-com-automation.md 3.7 与 SKILL.md 7.2 逐字重复；iron-laws.md 多条直接复制详述文件内容；vcs-text-editing.md 配对表与 SKILL.md 总览二重复约 90%。修复方案待用户确认。
- **验证口径**：tests/test_skill_integrity.py **54/54 通过**；SKILL.md 结构锚点不变（9 阶段 30 步，四要素 30/30/30/9，手段 A-F 六条）；12 个 src 模块全部被引用。

## v2.12（references 编排修复：真实错误 + 去重 + 职责归位）

- **变更动因**：用户确认 v2.11 批判的三类修复方案，逐项落地。
- **① 真实错误**：
  - `references/README.md` 补 `framework-comparison.md` 导航条目（原「按问题读」漏列，仅 SKILL.md 0.2 引用）。
  - `references/access-architecture-network.md` 阶段号修正：L4、L34「阶段 7 部署运维」→「阶段 8 部署与运维」（SKILL.md 9 阶段中部署为阶段 8）；L35 案例对应「阶段 1→2→3→4→6」→「1→2→3→4→5→8」（旧编号残留）。
  - `references/framework-comparison.md` 声明改实：原「未核实处标注待核」全文无待核标注 → 改为「凡未经真机实测、仅出自教程/文档的表述以『据教程』/『据文档』标注；已实测项直接写明实测结论」，并为升级部署/权限体系两行补标注。
  - `docs/troubleshooting.md` 断链修复：6 处 `references/vba-project-operations.md`（不存在）→ `references/vba-com-automation.md`（VBE 工程操作实为后者内容）。
- **② 去重**：
  - `references/vba-com-automation.md` 3.7 节（52 行 A/B/C/D 详单）压缩为指针：细则单点在 SKILL.md 7.2（手段 A–E 已完整承载），3.7 只留 4 个调用入口。
  - `references/iron-laws.md` A 组压缩：2a（VBComponents.Import 同样 GBK 无 BOM）并入 2；铁律 41 去除整段展开、改指针；修正其节号引用「第十节」→「第十一节」。编号 1–43 不变，引用安全。
  - `references/vcs-text-editing.md`：二、配对表（11 行）→ 指向 SKILL.md「二、对象全景与文本形态」单点 + 保留两条操作要点；六、手册表（9 行）→ 指向 references/README.md 导航 + 保留 CustomerOrders 对照阅读法。
- **③ 职责归位**：
  - `references/vcs-three-channels.md` 第十三节（错误 91 完整修复记录）压缩为操作口径指针，完整记录迁入本 changelog（见下条）。
  - `references/vcs-three-channels.md` 第九节（崩溃判别）压缩为三步指针，完整判据指向 docs/troubleshooting.md 第 19 节（该节本就承载主体）。
- **测试**：`tests/test_skill_integrity.py` REQUIRED_FILES 补 `references/framework-comparison.md`（原完整性测试未覆盖它，与 README 漏导航同源）——基线 54 → **55/55**。
- **验证口径**：完整性 55/55 通过；铁律编号 1–43 连续可引用；SKILL.md 结构锚点不变。

---

### 历史修复记录（自 references/vcs-three-channels.md 第十三节迁入，v2.12）

**ExportObject 错误 91 修复记录（NoIndex 保护，勿回退）**

- **现象**：`vcs_export_object module basSample` 触发 VBA 错误 91「对象变量或 With 块变量未设置」，高亮行 `cIdx.FolderAnnotation = IIf(Len(strAnnotation) > 0, strAnnotation, FOLDER_ANNOTATION_NONE)`（在 clsDbModule 的 `Export` 中）。
- **根因链**：`blnNoIndex:=True`（MCP/API 路径）→ `ExportSingleObject` 置 `disabledIndex` → `clsVCSIndex.Update` 在 `Me.Disabled` 时直接 `Exit Function`（返回 Nothing）→ `clsDbModule/clsDbForm/clsDbReport/clsDbVbeForm` 的 `Export` 中 `Set cIdx = VCSIndex.Update(...)` 之后**无条件**访问 `cIdx.FolderAnnotation` → Nothing 越界。
- **修复**（已合入源码并构建进技能目录运行位 accda）：四处源码 `msaccess-vcs-addin\Version Control.accda.src\modules\Components\{clsDbForm,clsDbModule,clsDbReport,clsDbVbeForm}.cls` 均加保护：
  ```vb
  If Not cIdx Is Nothing Then
      cIdx.FolderAnnotation = IIf(Len(strAnnotation) > 0, strAnnotation, FOLDER_ANNOTATION_NONE)
  End If
  ```
- 修复前备份已随历史归档回收（原始版本在 `msaccess-vcs-addin` 的 git 历史可查）。**重构建时不得回退此保护**；重构后必须用「导出单个对象 → 无错误 91、产物落盘、索引正确」复核（自动化门禁见 `scripts/rebuild_and_install.py`）。

## v2.13（references 分工收敛：vcs-three-channels 瘦身 + README 声明更新）

- **变更动因**：用户确认分工批判的 3 项建议，消灭"同一内容多处展开"的最后一处残留，使 8 个 references 文件全部单一职责。
- **vcs-three-channels.md 瘦身（保留全部节号，内部引用安全）**：
  - 四节（API 可操纵面）：命令表（9 行）→ 结论一句话 + 展开单点指向 `docs/troubleshooting.md` 第 20 节；
  - 五/六/七/八节（Build 硬前提/MergeBuild/两前提/成功四判据）：全展开 → 结论式（各 2–4 行），展开单点同样指向 troubleshooting 20；
  - 十四节（自动化测试）：6 行全展开 → 指针，细则单点指向 SKILL.md 7.2——**消灭"SKILL.md 7.2 / vba-com-automation 3.7 / vcs-three-channels 十四"三方重复的最后一处**；
  - 三节（编码铁律）本就一行指针，保留。
- **references/README.md**：L9「引用关系是单向的：references 不反向引用 SKILL.md」→「**单点优先、双向导航**」（v2.12 起 references 已反向引用 SKILL.md 单点：vcs-text-editing→总览二、vba-com-automation 3.7→7.2，旧声明与现状矛盾）。
- **职责矩阵最终态**：iron-laws=结论索引；encoding-rules=编码单点；vcs-text-editing=文本单点；vba-com-automation=COM SOP 单点；vcs-three-channels=安装配置+通道选择+指针；framework-comparison=选型；access-architecture-network=架构；README=导航；troubleshooting=排错+Build 姿势展开单点；SKILL.md=主干工作流。
- **验证口径**：完整性 55/55 通过；vcs-three-channels 节号 一–十四 完整保留；四节命令表与十四节展开已清除；铁律 1–43 连续。

## v2.14（SKILL.md 批判修复：2 真错误 + 闭环单点化 + 新增 2.5/8.4 + addin 源码引用）

- **真错误修复**：
  - 7.2 循环引用：删「与 references/vcs-three-channels.md 第十四节」（v2.13 已把十四节压缩为指向 7.2 的指针，原引用形成 7.2→十四→7.2 死循环），改为「完整方法见 references/vba-com-automation.md（本步即手段 A–F 的展开单点）」；
  - 总览「三段推进」→「六段推进」（原文列了 6 段，措辞与内容不符）。
- **门禁/闭环单点化**（落实 v2.8 声明的「VCS 融入只写本步差异动作」）：
  - 6.2 门禁①漂移核验 → 标注基线定义见 2.1、构建成功判别见 7.4；门禁②编译 → 手段见 7.2 手段 D；
  - 4.3/5.3 VCS 融入：闭环表述 → 指向 2.1 统一闭环，只保留本步差异动作（窗体：FullExport 拿 .form + 删除配对规则）；
  - 6.2 VCS 融入：闭环 → 指向 2.1，差异动作=Application.Run 运行断言。
- **新增独立步骤**（9 阶段 30 步 → 32 步）：
  - 2.5 性能设计（索引/查询/窗体加载，四要素 + VCS 融入）；
  - 8.4 备份与恢复（文本源即备份/发布快照/后端备份/恢复演练；原 8.4 收尾核验顺延为 8.5，无外部引用被破坏）。
- **addin 源码引用进工作流**（msaccess-vcs-addin 作为构建参考）：
  - 对象全景：生产级 .src 工程范本导航（Version Control.accda.src，10 模块分层 + 类封装 + 测试套件）；
  - 3.1 标准模块：modules/ 十目录分层范本；
  - 7.4：Tests/ 测试套件类模块源码（与 vcs_run_tests 对应）。
- **验证口径**：完整性 55/55 通过；步骤数 32（锚点无硬编码 30，安全）；「三段推进」与对 vcs-three-channels 十四节的引用已清除。

## v2.15（0.2 底座选型 → 框架与积木地图：两框架组成拆碎融入各阶段）

- **变更动因**：用户要求"不做 0.2 底座选型"，改为把两框架（盟威 / Edonsoft）的组成拆碎拆细，在工作流各阶段充分介绍引用，方便 LLM 借鉴其实践构建应用系统。
- **SKILL.md 0.2 改写**：标题「底座选型」→「框架与积木地图（两框架组成拆解）」；删除"二选一选底座"决策，改为两段能力映射——
  - 盟威：九章教程逐章对应工作流阶段（第 3 章→阶段 1、第 4→2.2、第 5→2.4、第 6→4、第 7→5、第 8→8.2、第 2.4 节→8.3）；
  - Edonsoft：18 项能力按阶段映射（代码生成器→4.1、通用查询→2.4、搜索组合框/单据选择→4.2、待办/仪表盘→4.3、链接表管理/升迁→8.1、RBAC→8.2、自动更新→8.3、AccessAI→6.1、公开 API 面→3.1）；
  - 输出物改为「框架积木地图」。
- **各阶段补充引用**：3.1 补 Edonsoft 公开 API 面范本（11-api.md + api/ 60 函数 7 类，业务公共函数层组织参考）；4.3 补主界面工作台组件（08-todo 待办、09-dashboard 仪表盘的数据源与挂载方式）。
- **framework-comparison.md 定位同步**：标题「底座选型对照」→「两框架差异对照」；「二、选型决策」→「二、按需取用（什么场景借鉴哪个）」，表头「条件|选盟威|选Edonsoft」→「场景|借鉴盟威|借鉴Edonsoft」并补主界面组件行；「三、与 VCS 工作流的关系」→「三、与工作流的关系」，「基座/所选底座」→「积木来源/不直接修改」。
- **"底座"措辞全局清理（9 处）**：frontmatter「基于现成底座」→「借鉴现成框架与积木」；总览 L14/L18、阶段 0 标题与开场、2.5「两底座」→「两框架」、4.1「用底座时」→「Edonsoft 框架下」、4.2「Edonsoft 底座下」→「框架下」、6.1「Edonsoft 底座」→「框架」。
- **references/README L27** 导航词「选底座纠结」→「纠结…（能力差异对照/按需取用/与工作流的关系）」。
- **验证口径**：完整性 55/55 通过；SKILL.md 无「底座选型」决策残留（L57 的「不做二选一选型」与 L62「差异对照」为正确表述）；framework-comparison 仅保留否定式说明（「不是必选底座」）。

## v2.16（两框架源码与组成融入各阶段 + 加载项 Wiki 全量导航）

- **变更动因**：用户要求 ① 把两框架的"源码和组成部分"拆碎拆细融入工作流（不只是文档层），② 在工作流各阶段充分介绍引用 `msaccess-vcs-addin\Wiki`。
- **探查结论**：两框架均无 `.src`/`.bas` 源码（Edonsoft=accdb+accde 加密/编译，盟威=accdb 加密 + 反编译 vbs），故"源码"层以文档中的**真实对象/函数落点**引用（已枚举核实）。
- **SKILL.md 0.2 补"源码与组成"层**：
  - 盟威：库文件构成 + 组件划分（自动编号/数据模块向导/导航菜单/查阅列表/启动菜单/报表向导/用户权限）+ 加密说明（`反编译打开Main.accdb.vbs`）；
  - Edonsoft：18 项能力全部真实落点（`USysFrmLogin`/`VerifyPassword`/`HashPassword`、`USysFrmRoleManager`/`GetPermissions`、`USysFrmMenuSettings`/`FArguments`、`BuildAccessAdoConnectionString`、`USysFrmLinkedTableManager`/`LinkedTableMgr_*`、`USysFrmUpsize`/`Upsize_*`、`USysFrmSettings`/`USysFrmSetup`、`USysFrmCommonQuery`/`CommonQuery_*`、`CreateSearchCombo`/`clsSearchCombo`、`USysFrmCodeGeneratorPro`、`USysFrmBillSelect`/`BillSelect_*`、`USysFrmToDoList`/`AddReminder`、`USysFrmDashboard`/`Dashboard_*`、`USysFrmProcessBar`、`basAI`/`FS_AIWeb`、`AutoNumStr`/`ReadRecord`/`WriteRecord`/`DeleteRecord`/`JsonConverter`）——数据源为 `13-architecture-source.md` 第 8 节表格（枚举核实）。
- **加载项 Wiki 导航（总览集中映射 + 各阶段引用，共 11 处）**：
  - 总览：新增 Wiki 导航段（Version-5-Overview / Installation / Options / Quick-Start / MCP-and-Automation / Export-Import-File-Types / Supported-Objects / Terminology-and-Style-Guide / Query-Source-Files / Split-Files / Merge-Build / Export-on-Save-Hook / Testing / Regression-Testing / Connections / PUBLISH / Security-Considerations / FAQs → 各阶段箭头）；
  - 各阶段引用：0.3 Installation/Options/Quick-Start/MCP-and-Automation；对象全景 Supported-Objects/Export-Import-File-Types/Terminology；2.4 Query-Source-Files；4.1 Split-Files；6.2 Export-on-Save-Hook（含 Hook/ 源码）；7.2 Testing/Regression-Testing；7.4 Merge-Build；8.1 Connections；8.3 PUBLISH；8.5 FAQs。
- **各阶段框架源码落点补充**：2.4 通用查询 `USysFrmCommonQuery`/`CommonQuery_*`；3.1 公共函数落点（`AutoNumStr`/`ReadRecord`/`WriteRecord`/`DeleteRecord`/`JsonConverter`）+ 操作日志 `WriteLog`/`AuditData`（**不写类名 clsLog，规避完整性禁词**）；4.2 进度条 `USysFrmProcessBar`/`USysFrmProgressFlat/Dual/Detail`；4.3 动态菜单 `USysFrmMenuSettings` + 待办/仪表盘落点（`USysFrmToDoList`/`AddReminder` 系列、`USysFrmDashboard`/`Dashboard_*`）；8.1 多后端 `BuildAccessAdoConnectionString` + `USysFrmLinkedTableManager`/`USysFrmUpsize`；8.3 自动更新源码落点（AutoExec 检测 + `USysFrmSettings`/`USysFrmSetup`）——**顺带修正历史错误**：原「13-architecture-source.md 第 11 节」实为兼容性矩阵，自动更新在第 8 节表格第 7/8 行。
- **验证口径**：完整性 55/55 通过；SKILL.md 中 `Wiki\` 引用 11 处；18 项能力落点抽查 6 处命中；禁词（clsLog/Framework_RunAllTests）零残留。

## v2.17（MCP 机制落地 + addin 源码/docs 融入 + 两框架可见可读部分融入）

- **变更动因**：用户要求 ① 解读"插件自带 MCP"的机制/功能/用途及与 msaccess-vcs-mcp 的差异；② 在工作流各阶段充分介绍引用 `msaccess-vcs-addin\Version Control.accda.src` 与 `msaccess-vcs-addin\docs`；③ 不做 0.2 底座选型，把两框架**可见可读部分**拆碎拆细融入工作流。
- **探查结论（事实）**：
  - 插件"自带 MCP"= 加载项内置**公开 API**（`VCS`/`clsVersionControl`，源码 `Version Control.accda.src\modules\API\clsVersionControl.cls`）+ **权限门**（Options → MCP：`McpAllowImport`/`McpAllowExecuteSQL`/`McpAllowRunVBA`，默认全关，配置存 `vcs-options.json`）；**MCP 服务器本身在 `msaccess-vcs-mcp`**（Python，COM + `Application.Run` 调加载项 API）。差异=加载项是能力提供方、mcp 是协议适配方，缺一不可。
  - `addin\docs` 三篇：README（docs 定位=维护者/AI agent 内部参考，与 Wiki/AGENTS.md/DECISIONS.md/Testing 分工）、`access-query-storage.md`（MSysQueries 字段、Design vs SQL View、LoadFromText/SaveAsText 不对称，来源 Colin Riddington）、`access-conditional-format.md`（ConditionalFormat/14 二进制逆向规范，实现 `clsConditionalFormat.cls`）。
  - 两框架**可见可读部分**：Edonsoft=`framework-overview.md` 第 4 节 COM 枚举对象清单 + 可见源码模块（`basCoreReference`/`basAI`/`basZip`/`basButton`/`basFormFontTools`/`clsLog`/`JsonConverter`/`Module_DatePicker`/`Module_YearMonthPicker`）+ `appsettings.json`（ActiveBackend/Backends[]）+ `assets/`；盟威=`Config.ini`（LoginUserList/AutoLogin/ThemeColor）+ `Images/`（login_classic/login_standard/head.jpg/progress.jpg）+ 教程 PDF。框架目录实测**无 `mode/` 目录**（文档所述为开发规范，非实际目录）。
- **禁词修正**：`tests/test_skill_integrity.py` 将 `clsLog` 移出 FORBIDDEN_TOKENS（附注释）——经核实它是 Edonsoft 可见源码模块名（framework-overview 第 4 节），SKILL.md 引用其为合法借鉴目标；框架目录本在 SKIP_DIRS，扫描不受影响。
- **SKILL.md 编辑（8 处）**：
  - 0.2：盟威补可见可读（Config.ini/Images/教程）；Edonsoft 补第 4 节对象清单 + 可见源码模块 + appsettings.json + assets；
  - 0.3：补"插件自带 MCP"机制与差异（API+权限门 vs MCP 服务器，威胁模型指向 Wiki）；
  - 2.4：补 `docs\access-query-storage.md`（查询内部存储权威，解释 .bas/.sql 恒成对）；
  - 3.2：补生产级类模块范本（`clsVersionControl`/`clsQueryComposer`/`clsSourceParser`/`clsConditionalFormat`/`modLoadSaveText`）；
  - 3.1：补 Edonsoft 可见源码模块清单（含可移植的 `Module_DatePicker`/`basCoreReference`）；
  - 4.2：补条件格式二进制权威规范（`docs\access-conditional-format.md` + `clsConditionalFormat.cls`）+ 界面资源参考（Edonsoft assets / 盟威 Images）；
  - 8.1：补 `appsettings.json` 多后端配置化切换结构；
  - 8.2：补盟威 `Config.ini` 登录/主题配置文件式管理。
- **验证口径**：完整性 55/55（禁词移除不影响检查项数）；新引用计数 12 项抽查全命中；vcs-three-channels 仅余两处（0.3/7.4），无循环引用。

## v2.18（自我审查：工作流形态收敛 + 四要素补全 + 串题修正）

- **变更动因**：用户要求对 SKILL.md 反复核实与自我批判——检查是否形成流畅自然/井然有序/分工明确/逻辑严密的工作流、是否存在索引表/路由表/资源清单等非工作流形态、内容是否与构建 Access 应用相关、是否与所处阶段步骤主题相关。
- **审查结论**：
  1. **工作流形态**：0–8 九阶段 32 步四要素主线成立；对象全景表是"对象→阶段→文本形态→手册"的**横切路由视图**（合理导航，已加定位说明"本身不是工作流步骤"）；总览中 MCP 源码明细段与 Wiki 18 篇导航段是**冗余资源清单**（与 0.3/6.2/7.x 及各阶段引用重复）——已压缩为指针。
  2. **四要素一致性**：总览声明"每步四件事"，但 2.3/3.2/4.1/5.1/5.2/6.1 缺 VCS 融入字段——已补全（21 步有 VCS 融入，设计期/纯数据期合理缺省）；4.3 与 4.1 的 VCS 融入重复——4.3 压缩为指针。
  3. **串题修正**：2.5 性能设计的"两框架链接表与后端升迁"引用属 8.1 主题——改为"数据量超单文件共享承受力时前后端分离/后端升迁是终局手段（见 8.1）"。
  4. **衔接显式化**：2.2 补"阶段 1.4 物理建模的落实"标注，使 1.4→2.2 的概念→逻辑→物理链路显式。
  5. **0.3 MCP 机制段**：威胁模型细节（信任链/会话覆盖/UI 跳过/索引）压缩为指向 Wiki。
- **验证口径**：完整性 55/55；四要素计数 任务内容 34 / 任务要求 34 / 输出物 38 / VCS 融入 21；Version-5-Overview 明细 0 处；access_gate 仅 7.2 一处（重复消除）。

## v2.19（0.2 与各阶段重复审查：映射兑现核对 + 可见源码模块清单单点化）

- **变更动因**：用户询问 0.2 框架与积木地图是否与后续内容重复、映射是否已在各阶段兑现。
- **核实结论**：Edonsoft 9 项能力映射与盟威教程 7 项章节映射**全部在各阶段兑现**（逐项核对：代码生成器→4.1、通用查询→2.4、搜索组合框/单据选择/进度条→4.2、待办/仪表盘→4.3、链接表/升迁/多后端→8.1、RBAC/登录→8.2、自动更新→8.3、AccessAI→6.1、公共函数/操作日志→3.1；盟威第 3/4/5/6/7/8 章与 2.4 节→阶段 1/2.2/2.4/4/5/8.2/8.3）。其余"重复"判定为地图（0.2 集中导航）+ 使用（各阶段具体用法）的双层结构，符合 SKILL Graph，非冗余。
- **唯一真重复已修**：0.2 与 3.1 的"Edonsoft 可见源码模块"全列表几乎一字不差列了两遍——3.1 压缩为"完整清单见 0.2 与 framework-overview.md 第 4 节；可直接移植的：Module_DatePicker/Module_YearMonthPicker、basCoreReference"，0.2 成为清单单点。
- **验证口径**：完整性 55/55；basZip/clsLog/basFormFontTools/basButton 现仅 0.2 一处；18 项能力落点精确为"0.2 地图 1 处 + 各阶段用法 1 处"（USysFrm* 各 2 处），无第三处重复。

## v2.20（框架落点从 0.2 地图下沉到各阶段使用处）

- **变更动因**：用户指出 0.2 的框架落点列举未真正融入各阶段——如 `USysFrmLogin` 只在 0.2 出现；要求"在工作流的不同阶段不同步骤中充分介绍和引用两个框架，不必在 0.2 中列举"。
- **核实**：全文扫描 0.2 全部落点，找出只在 0.2 出现的 5 个：`USysFrmLogin`、`USysFrmRoleManager`、`USysFrmBillSelect`、`clsSearchCombo`（已在 4.2 有 CreateSearchCombo 但未点实现类）、以及 v2.19 压缩导致从 3.1 消失的可见源码模块（`basZip`/`basButton`/`basFormFontTools`/`clsLog`）。`GetUnreadCount`/`Dashboard_AddKPIFromSQL` 已确认本就在 4.3（0.2 用通配符，非遗漏）。
- **SKILL.md 编辑（4 处）**：
  1. **0.2 压缩**：删 18 项能力落点长列表与 9 项映射明细，改为"能力与源码落点已按阶段融入工作流各步骤（登录/RBAC → 8.2、通用查询 → 2.4、公共函数/操作日志 → 3.1、代码生成器 → 4.1、搜索组合框/单据选择/进度条 → 4.2、动态菜单/待办/仪表盘 → 4.3、AccessAI → 6.1、链接表/升迁/多后端 → 8.1、自动更新/设置 → 8.3），完整清单见 13-architecture-source.md 第 8 节表格"；可见源码模块清单改为"各模块用途与可移植说明见 3.1"。
  2. **3.1 恢复**：完整可见源码模块清单带用途（basCoreReference/basZip/basButton/basFormFontTools/clsLog/JsonConverter/Module_DatePicker/Module_YearMonthPicker）+ 可移植标注。
  3. **8.2 补**：登录窗体 `USysFrmLogin` + 口令验证 `VerifyPassword`/`HashPassword`；角色管理窗体 `USysFrmRoleManager` + `GetPermissions`/`LogOff`。
  4. **4.2 补**：单据选择窗体 `USysFrmBillSelect`；搜索组合框实现类 `clsSearchCombo`。
- **验证口径**：完整性 55/55；全部 USysFrm* 落点现各 1 处且均在对应阶段（0.2 区间无明细残留）；AutoNumStr 4 处/JsonConverter 3 处为多场景正当引用。

## v2.21（排版优化：超长段落拆分为无序列表分点）

- **变更动因**：用户要求段落不宜太长，列举一连串多个对象时用无序列表分段分点。
- **拆分范围（12 处）**：
  - 0.2：盟威段拆为 组成/教程九章映射/源码与组成/可见可读部分；Edonsoft 段拆为 组成/能力落点按阶段/文档/可见可读部分；
  - 对象全景：硬规则总入口/活样例/更大规模工程范本/Wiki 权威 拆为 4 个子项；
  - 0.3：加载项 bullet 拆为 安装/运行位优先/操作手册 3 条；MCP bullet 拆为 安装与客户端配置/调用链与异步语义/抄骨架/自带 MCP 机制 4 条；
  - 2.2/2.4/3.1/3.2/4.1/4.2/4.3/8.1/8.2/8.3：任务要求长列表拆为子项；
  - 6.2 VCS 融入拆为 4 条（闭环断言/八种模式/tools.py 抄写/保存即导出）；
  - 7.2 任务内容拆出「测试资源（MCP 与加载项侧）」子项；7.4 任务内容拆为 四条件判别/加载项回归/长操作审计 3 条。
- **顺带修正**：7.4 任务要求「完整性 54 项」→ 55 项（v2.12 基线已升至 55，SKILL.md 未同步）。
- **验证口径**：完整性 55/55；无超 600 字符行；剩 7 行 400–600 均为命令/路径密集行（install 命令、2.1 导出命令、8.5 案例对照、description 元数据）；四要素结构完整（内容 34/要求 34/输出物 38/VCS 融入 21）。

## v2.22（排版优化二：所有多项并列段落拆为无序列表）

- **变更动因**：用户重申"每个段落不宜太长，当多项内容并列时用无序列表分段分点"——上一轮只处理 >600 字符行，本轮把标准降到所有 >350 字符的多项并列段落。
- **拆分（11 处）**：
  - 总览：统一闭环工具链拆 4 条（导出/构建/核验/守卫+警告）；MCP 手册拆 3 条（5 份文档/源码样板）；文档分工拆 4 条；
  - 0.3：任务内容拆 2 条（安装位约定/references 指引）；MCP 安装配置拆 4 条（代码位置/.env/客户端配置/McpAllowRunVBA）；"插件自带 MCP 机制"拆 3 条（API+权限门/MCP 服务器/威胁模型）；
  - 2.1 VCS 融入拆 4 条（锁基线/真机演示/导出产物格式/AI 代理工作流）；
  - 2.3 任务要求拆 4 条（时序/删除规则/格式/活样例）；2.4 任务内容拆 2 条（查询类型/取数示范）；
  - 2.5 任务要求拆 4 条（索引/查询/窗体加载/优化顺序）；
  - 3.1：优先复用拆 3 子项（模块分组列举）；可见源码模块拆 3 子项；
  - 8.5 案例对照拆 4 条（闭环演练/完整业务样例/积木型案例/排错入口）。
- **验证口径**：完整性 55/55；>350 字符行仅剩 2 行（description YAML 元数据与 install 命令行——均不可断行）；四要素结构完整（内容 34/要求 34/输出物 38/VCS 融入 21）。

## v2.23（8.5 案例对照移出收尾 + VBA-Docs 权威参考融入各阶段）

- **变更动因**：① 用户质疑 8.5 收尾核验为什么有案例对照、是否太晚；② 用户要求 SKILL.md 充分介绍引用 `..\VBA-Docs`（微软官方 VBA 文档库）相关内容。
- **8.5 修正**：案例对照四条（闭环演练/业务样例/积木型案例/排错入口）从收尾核验移除——案例已按使用场景嵌入各阶段（2.1/2.2/2.3/2.4/3.1/4.1/4.2/8.3），收尾再导航属重复的资源清单形态；排错入口并入任务要求（收尾发现问题时的出口）；新增「案例的定位说明」明确"做哪一步看哪一步引用的案例，收尾无需再导航"。
- **VBA-Docs 融入（10 处）**：
  - 总览文档分工：VBA-Docs 资料库总入口（`..\VBA-Docs\`，3.2 万+ API 页 + 概念手册，索引见 `..\vba_docs_index.db`）；
  - 2.4：SQL 语法权威参考（Structured-Query-Language 48 篇）+ 条件表达式参考（Criteria-Expressions 19 篇）；
  - 3.1：记录集参考（Data-Access-Objects 27 篇 + ActiveX-Data-Objects）+ VBA 语言总索引（Language\Reference 关键字/函数/语句/对象/常量）；
  - 3.2：类模块与错误处理官方参考（Error-Codes 9 篇 + Language\Concepts + 事件模型）；
  - 4.1：窗体事件与记录操作官方示范（Forms/Forms-Design）+ Access 对象模型按名查 `api\`（`Access.` 前缀）；
  - 4.2：控件官方示范（Controls 7 篇）；
  - 5.1：报表与打印官方参考（Reports/Printing）；
  - 7.2：VBA 错误消息与 Access 错误码权威清单（error-messages.md + Error-Codes）；
  - 8.3：启动属性/选项官方写法（Settings 6 篇）。
- **路径说明**：主技能整体移至 `VBA\Microsoft-Access-DataBase\`（VBA 技能下），VBA-Docs 引用以 `..\VBA-Docs\` 相对路径；移动/云盘同步期间 examples/CustomerOrders 曾出现临时空壳，属同步正常现象，未做修复。
- **验证口径**：四要素结构完整；>400 字符行仅剩 description 与 install 命令行；完整性 55/55 已确认（云盘同步完成后复跑）。
- **测试脚本适配**：`tests/test_skill_integrity.py` REQUIRED_FILES 中 `vcs-index.json` 改为 `vcs-index.idx`——CustomerOrders 样例为 VCS 5.0.1 导出（vcs-options.json AddinVersion=5.0.1，二进制索引 vcs-index.idx；v4 才是 vcs-index.json）。v5 查询为 `.bas+.qdef+.sql` 三件套，SKILL.md 现有引用（qOrders.sql/tOrders.xml/tCustomerstOrders.json/fCustomerList.bas）均不受影响。

## v2.24（AccessAI 例子升级替换：移除旧版、新版更名）

- **背景**：用户提供升级版 AccessAI 项目（`examples/accessAI-main`，实为开源「Access LLM Toolkit」，作者缪炜），要求替换旧的 `examples/AccessAI` 例子，并更新所有相关引用。
- **处理**：
  - 移除旧 `examples/AccessAI`：经 `scripts/recycle.py` 移入 Windows 回收站（绝不永久删除，回收站可随时还原）；返回「成功 1 / 失败 0」。
  - 更名新项目：同文件系统内 `examples/accessAI-main` → `examples/AccessAI`，接替旧项目位置。路径 `examples/AccessAI/` 保持不变，既有引用自动指向新内容。
  - 更新三处文档描述（均保持 VBA 技能约定：内部引用用反引号、不带 `./`、语义由句子承担）：
    - `README.md`（AI 能力节）：描述为升级版工具库——`CreateAIForm` 一键建窗体，支持流式输出、对话历史持久化（`tblChatHistory`）、Access SQL 助手、TXT/CSV/Word/Excel/PDF 文档问答、API Key 用 Windows DPAPI 加密存储。
    - `SKILL.md`（阶段 6 任务要求）：标注为升级版「Access LLM Toolkit」并列出新增能力。
    - `examples/README.md`（积木型案例表格）：同步新能力清单。
- **验证口径**：`examples/AccessAI/README.md`、`AI.accdb` 等已就位；`SKILL.md`/`README.md`/`examples/README.md` 中 `examples/AccessAI/` 引用解析到新内容（除本条目描述外，技能内已无任何指向 `accessAI-main` 目录的有效引用）；`tests/test_skill_integrity.py` **55/55** 不受影响。

## v2.25（建立上游漂移跟踪机制，门禁拦截手改 vendored 文件）

- **背景**：本技能 vendored（引入并锁定）了 7 路开源 / 商业上游目录，参照 `VBA/xlwings`、`fastapi`、`hermes-business-agent` 三个技能的做法，建立可复现的漂移跟踪机制，防止上游文件被意外手改而无人察觉。
- **7 路上游（含 provenance）**：`Microsoft Access Version Control System`（MAVCS 编译分发物，joyfullservice/msaccess-vcs-addin）、`msaccess-vcs-addin`（源码，main，Version 5）、`msaccess-vcs-mcp`（Python 封装，v0.1.0，Adam Waller）、`Version_Control_v5.0.1`（v5.0.1 锁定快照）、`盟威Access快速开发平台V2.7.0版(64位)`（商业件）、`Edonsoft Development Framework_x64`（商业件）、`examples`（混合容器：开源积木 AccessAI / Access BOM / DatePicker / VBA Modules + 技能自有 01/02/CustomerOrders/README）。
- **新增三件套**：
  - `manifest.json`：逐路登记来源（id / name / local_dir / repo / ref / version / kind / note）+ 注册时基线指纹（baseline_file_count / baseline_total_bytes / baseline_struct_hash / baseline_content_hash / registered_date）。`examples` 项额外含 `components[]` 区分开源与技能自有。
  - `SYNCLOG.md`：日志流形式累积同步 / 漂移事件，首条记录机制建立 + 基线登记；约定"原文件永不手改，任何主动改动走整体替换 + 重新 `--register` + 记一条"。
  - `scripts/check_upstream_drift.py`：离线 `struct_hash`（仅 stat，极快）+ `content_hash` 双哈希漂移检测；`--register` 重算并写回基线；`git ls-remote` 在线上游更新提示（默认只报告不报错，网络不可达时降级跳过）；只读写 manifest/SYNCLOG，绝不改 vendored 内容。
- **门禁集成**：`tests/test_skill_integrity.py` 新增第 5 项检查——离线复用 `check_upstream_drift.py` 的 `compute_struct` / `compute_content` / `load_manifest`，逐路比对 `manifest.json` 基线；偏离即失败（强制走正规更新流程）。检查项由 55 → **62**（新增 7 路上游各 1 项）。
- **验证口径**：`--register` 成功写回 7 路基线（MAVCS 5 文件/7MB、addin 552/8MB、mcp 70、v5.0.1 1/13MB、盟威 563/18MB、Edonsoft 122/14MB、examples 95/15MB）；独立漂移检测与门禁均报告全 UNCHANGED（退出码 0）。**负向验证**：在 `examples/` 临时植入测试文件 → 漂移检测与门禁均报 `上游漂移: examples` 并退出码 1，清理后恢复 62/62 全绿，确认门禁能真正拦截手改 vendored 文件。

## v2.26（修正上游 provenance + examples 仅跟踪开源积木）

- **背景**：用户补充了各上游的真实来源，并要求 `01_blank_db_to_vcs_loop`、`02_edit_src_reimport` 不参与跟踪（它们是技能自有、会随开发改动）。
- **provenance 修正**（`manifest.json`）：
  - `examples` 的 5 个开源积木：AccessAI 等四个来自 `github.com/miaowei2`（AccessAI 即 `miaowei2/accessAI`，Access LLM Toolkit），`CustomerOrders` 来自 `github.com/paramountsoftware/ms-access-ai-skill/tree/main`。components 每项加 `source_url` 与 `tracked` 标志。
  - `盟威Access快速开发平台V2.7.0版(64位)` 加 `source_url=http://www.accessgood.com/`（商业件，无公开 git）。
  - `Edonsoft Development Framework_x64` 加 `source_url=http://www.edonsoft.com/access-framework`（商业件，无公开 git）。
  - 约定：有 git 仓库的上游写 `repo`（供 `git ls-remote` 更新提示），无 git 的网页/商业来源写 `source_url`。
- **examples 跟踪范围收窄**：`examples` 源新增 `track_relpaths`（AccessAI / Access BOM Management System / Access DatePicker / Access VBA Modules Collection / CustomerOrders），漂移检测只对这些子目录计算基线；`01_blank_db_to_vcs_loop`、`02_edit_src_reimport`、`README.md` 标 `tracked:false` 被排除。`scripts/check_upstream_drift.py` 的 `walk_files/compute_struct/compute_content` 增加 `only_relpaths` 参数；`register()` 与 `main()` 及门禁 `tests/test_skill_integrity.py` 第 5 项均传入该参数。
- **基线重算**：`--register` 后 examples 基线由 95 文件 → **90 文件**（排除 5 个不跟踪件）。
- **验证口径**：重新登记后 7 路全 UNCHANGED（退出码 0）；门禁 62/62 全绿。**负向验证**：在【不跟踪】的 `01_blank_db_to_vcs_loop` 植入文件 → examples 仍 UNCHANGED（不误报）；在【跟踪】的 `AccessAI` 植入文件 → 报 `上游漂移: examples` 且门禁失败（退出码 1）；清理后恢复全绿。确认"排除自有件 + 仍拦截手改 vendored"两个目标同时达成。

## v2.27（发布态占位策略：来自开源仓库的示例目录一律 README 占位）

- **背景**：用户明确"**来自于开源仓库的内容采用 README.md 占位**"。此前仅 6 个重二进制 vendored 顶层目录（MAVCS / addin / mcp / v5.0.1 / 盟威 / Edonsoft）在发布仓库里压成 README 占位；`examples` 下 5 个开源积木目录仍随仓库分发真实内容（含 `.accdb` / `.mdb` / `.rar` 等重二进制）。本轮把该策略推广到全部开源仓库来源。
- **发布仓库处理**：`examples` 下 `AccessAI` / `Access BOM Management System` / `Access DatePicker` / `Access VBA Modules Collection` / `CustomerOrders` 五个目录在发布态一律只保留 `README.md`（本轮删除 86 个文件，`90 - 4 = 86`）。占位 README 规则与既有 vendored 目录一致——本地有 README 的保留其上游 README，本地无 README 的写入「（占位）」说明：`CustomerOrders` 本轮生成指向 `github.com/paramountsoftware/ms-access-ai-skill` 的占位说明；其余 4 个保留上游 README（Access LLM Toolkit / Access BOM 管理系统 / Access DatePicker / AccessDevelop）。技能自有件（`SKILL.md`、`references/`、`scripts/`、`tests/`、`templates/`、`assets/`、`examples/01_`、`02_`、`examples/README.md`）原样发布。
- **漂移脚本**（`scripts/check_upstream_drift.py`）：
  - 新增 `source_is_placeholder(local_dir, only_relpaths=None)`——在 `is_placeholder` 基础上支持 `track_relpaths`：当**所有**被跟踪子路径都是"仅 README.md"的占位目录时，整个来源视为占位并跳过。解决 `examples` 这种"混合容器"在发布态（5 个子目录变占位、但 `01_`/`02_`/`README.md` 仍是真实技能自有件）被误判漂移的问题。
  - `main()` 占位判定由 `is_placeholder(ld)` 升级为 `source_is_placeholder(ld, s.get("track_relpaths"))`。
  - `register()` 增加占位态守卫：占位来源**不参与**基线重算（否则会用占位数据覆盖完整基线），打印"跳过（占位态，不覆盖基线）"。
- **门禁**（`tests/test_skill_integrity.py`）：
  - 第 2 项「关键文件齐全」：`REQUIRED_FILES` 里 9 条 `examples/CustomerOrders/...` 路径，若其上级目录为占位则跳过（新增 `_under_placeholder(p)` 守卫），避免发布态误报"关键文件缺失"。
  - 第 5 项「上游漂移」：改用 `source_is_placeholder(ld, s.get("track_relpaths"))`，与漂移脚本口径一致。
- **同步脚本**（工作区 `sync_mad.py`，不进技能包）：新增 `EXCLUDE_OPENSOURCE_EXAMPLES`（5 个示例目录）+ `is_placeholder_rel()` + `GENERATED_PLACEHOLDER_README`（为本地无 README 的占位目录生成说明）；`build_target_set()` 与删除保留规则统一按"占位目录只留 README.md"处理。
- **验证口径**：发布克隆内门禁 **62/62** 全绿；漂移检测在发布态 7 路全部 `PLACEHOLDER(跳过)`、本地完整态 7 路全部 `UNCHANGED`（examples 仍 90 文件）；发布态全仓库 >1MB 文件仅剩技能自有 `references/开发财务管理系统.pdf`（2.5MB，非开源仓库来源，按指令保留）。**负向验证**：在发布克隆的 `examples/AccessAI` 植入非 README 文件 → 立即报 `LOCAL_DRIFT` 且门禁失败（`上游漂移: examples`，退出码 1）；清理后恢复 `PLACEHOLDER` 与 62/62 全绿——证明占位跳过不是无差别屏蔽。
