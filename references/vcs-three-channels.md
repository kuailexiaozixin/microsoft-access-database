# 版本控制通道（VCS 加载项 / MCP / 脚本驱动）

> Access 的 `.accdb` 是二进制文件，不适合直接进 Git。Version Control System 加载项把库里的对象（表/查询/窗体/报表/模块/宏）导出成文本源文件，就能像普通代码一样做版本管理与协作。
> 本技能所有 VCS 工具**同源**：MCP 与驱动脚本底层都调用同一个加载项引擎。选哪条通道取决于"谁来操作"（人手点 Ribbon / 代码调 API / AI 代理走 MCP）。
> **版本控制不是收尾才做的单独一步，而是贯穿开发全生命周期的方法**：库一建好就导出 `.src` 作为单一真相源，之后每个阶段（建模、代码库、窗体、AI 改库、运维）的变更都走「改 `.src` 文本 → 构建/回写回库 → 漂移核验」的闭环。各阶段怎么用，见 `SKILL.md`；改文本的硬规则见 `references/vcs-text-editing.md`。

## 一、工具形态与安装（四个目录实体 + 技能目录自包含运行位）

技能目录里有四个 VCS 相关实体，关系如下（**不是"安装后形成"的先后关系，而是四种分发形态**）：

| 实体 | 形态 | 是什么 | 何时用 |
|---|---|---|---|
| `msaccess-vcs-addin/` | 上游**源码仓库** | 加载项的全部源码 + 文档 + `Ribbon\Build\MSAccessVCSLib_win64.dll`（官方编译的 Ribbon COM 加载项）+ `Ribbon\Ribbon.xml` | 看实现、自构建（`vcs_drive.py Build`）、读官方文档/AGENTS.md；Ribbon DLL 与 Ribbon.xml 从这里取 |
| `Microsoft Access Version Control System/` | **安装位 = 运行位（技能目录内，自包含）** | 运行所需的全部组件：`Version Control.accda`（修复版本体）+ `MSAccessVCSLib_win64.dll`（twinBASIC COM Ribbon 加载项）+ `Ribbon.xml`/`Ribbon.json`（功能区定义）+ `Worker.vbs`（后台启动器） | **所有操纵（Ribbon / API / MCP）唯一指向这里**；Menu Add-Ins、受信任位置、CLSID 均指向此目录 |
| `Version_Control_v5.0.1/` | **v5.0.1 官方安装程序** | 从 Releases 下载的 5.0.1 安装向导 `Version Control.accda`（14,274,560B，含内嵌资源）；`/cmd INSTALL` 时作为源 | 安装源留档；代码化安装的源输入 |
| `msaccess-vcs-mcp/` | MCP 服务源码仓库 | 把加载项 API 包装成 `vcs_*` 工具的 Python 项目；可编辑安装进 Python，运行入口 `python -m msaccess_vcs_mcp` | AI 编码代理自动驱动 |

> **自包含原则**：所有组件与功能都放在技能目录内运行（`Microsoft Access Version Control System\` 是加载项唯一的安装位兼运行位，MCP 源码与配置也在技能目录内），**不安装、不注册、不复制到技能目录之外**。唯一的例外是 COM 机制必需的注册表项——它们只写 `HKCU`（当前用户），且值全部指向技能目录内文件。历史旧安装位 `%AppData%\MSAccessVCS\` 已回收清理（移入回收站可还原），不再参与运行。

### 1.1 加载项安装与配置（全代码化，禁止手动弹窗）

1. 前提：Office 与 Access 位数匹配（x64 用 `MSAccessVCSLib_win64.dll`）；Access 已启用「信任对 VBA 工程对象模型的访问」。
2. **代码化安装**（等价于官方安装向导的全部步骤，无任何手动操作）：

```powershell
python scripts/install_vcs_addin.py `
  --source "msaccess-vcs-addin\Version Control.accda" `   # 构建好的 accda（安装源）
  --folder "Microsoft Access Version Control System" `    # 安装位=运行位（技能目录内）
  --ribbon-dll "msaccess-vcs-addin\Ribbon\Build\MSAccessVCSLib_win64.dll" `
  --ribbon-xml "msaccess-vcs-addin\Ribbon\Ribbon.xml" `
  --ribbon-json "Microsoft Access Version Control System\Ribbon.json"
```

该脚本按官方 `modInstall.InstallVCSAddin` + `modCOMAddIn.VerifyComAddIn` + `DllRegistration.DllRegisterServer` 的**实际执行代码**复刻：
   1. 写安装设置（`HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install\`：Install Folder=技能目录运行位、Use Ribbon=1）；
   2. 启动弹窗守卫 `scripts/fwguard.py`；
   3. 以 `MSACCESS.EXE <源accda> /cmd INSTALL` 触发官方安装代码（`UpdateAddInFile` 复制 accda 到运行位、`RegisterMenuItem` 写 Menu Add-Ins 三键 `&VCS Open`/`&VCS Options`/`&VCS Export All Source`、写受信任位置）；
   4. 部署 Ribbon COM 加载项：复制 `MSAccessVCSLib_win64.dll`/`Ribbon.xml`/`Ribbon.json` 到运行位，写 `HKCU\Software\Microsoft\Office\Access\Addins\MSAccessVCSLib.AddInRibbon`（FriendlyName/Description/ProgID/LoadBehavior=3 自动加载）；
   5. 六项自检（目标 accda、Menu Add-Ins 三键、受信任位置、Ribbon DLL、COM Add-Ins、自包含清单），全部通过才算装好。
3. 验证：`scripts/probe_vcs_api.py` 阶梯探针（`GetVCSVersion` 等逐命令试），返回版本号即通道可用；重启 Access 后「Version Control」Ribbon 选项卡出现。
4. 若只需只读导出审查：`ACCESS_VCS_DISABLE_WRITES=true` 或只调 `Export` 系命令。

### 1.2 MCP 安装与配置（全代码化）

1. 前提：Python 3.10+；加载项已按 1.1 装好；Access 已注册 COM。
2. 安装：`cd msaccess-vcs-mcp && pip install -e .`（可编辑安装，包位置即技能目录源码，随技能自包含）。
3. 配置：`msaccess-vcs-mcp\.env` 设 `ACCESS_VCS_ADDIN_PATH=<技能目录>\Microsoft Access Version Control System\Version Control.accda`（**必须指向技能目录运行位**）；向 MCP 客户端配置文件加入服务（Cursor 项目级 `.cursor/mcp.json`、用户级 `~/.cursor/mcp.json`；Claude Code 项目级 `.mcp.json`、用户级 `~/.claude.json`）：

```json
{ "mcpServers": { "msaccess-vcs-mcp": { "command": "D:\\Python\\cpython-3.13-windows-x86_64-none\\python.exe", "args": ["-m", "msaccess_vcs_mcp"] } } }
```

4. 端到端验证：`python tests/vcs_mcp_verify.py --db <宿主库.accdb>`（initialize → tools/list → `vcs_list_objects` → `vcs_call_vba` → `vcs_compile_vba` → `vcs_export_object`），全程无弹窗、返回 JSON、退出码 0 才算通过。
5. 跑代理生成的代码用 `vcs_run_vba` 需先在加载项选项开启 `McpAllowRunVBA`（一次性设置，用脚本写注册表 `HKCU:\Software\VB and VBA Program Settings\MSAccessVCS\Add-in` 代设，后续无需再碰）。

**环境变量分工（勿混淆）**：

| 变量 | 指向/含义 | 本技能是否设置 | 备注 |
|---|---|---|---|
| `ACCESS_VCS_ADDIN_PATH` | 加载项**运行位**文件（技能目录 accda） | ✅ 必设（`msaccess-vcs-mcp\.env`） | MCP 启动时加载哪个 add-in；缺失/指错则全部 `vcs_*` 工具失败 |
| `ACCESS_VCS_DATABASE` | **目标库**快捷方式（官方可选配置） | 可选 | 工具几乎都要求 `database_path` 参数；此变量仅为"不传参时的默认库" |
| `ACCESS_VCS_DISABLE_WRITES` | 只读审查开关 | 按需 | 生产库先只导出审查时设 `true` |
| `ACCESS_VCS_ENABLE_LOGGING` | 结构化 JSONL 日志 | 按需 | 排查 MCP 调用时序时设 `true` |

## 二、三种操作通道怎么选

- **Ribbon 选项卡（COM 加载项）**：人在 Access 里点「Version Control → Export / Build / Merge」。由 `MSAccessVCSLib_win64.dll`（twinBASIC COM 加载项，运行位=技能目录）提供，LoadBehavior=3 自动加载。适合人工巡检；自动化脚本不依赖它。
- **代码驱动（`Application.Run` / `scripts/vcs_drive.py`）**：主线程直接调加载项 API。适合确定性自动化——本技能所有出入库动作的默认通道，绕开 MCP 的异步/线程坑。
- **MCP（`vcs_*` 工具）**：AI 编码代理（Cursor、Claude Code 等）自动驱动。适合"AI 改文本源 → 构建回库"的循环；多做了路径校验、异步进度、权限开关。

AI 文本化开发的标准范式见 `references/vcs-text-editing.md` 与本技能 `SKILL.md` 阶段 6.2：导出文本 → 在文本上编辑 → 构建回库。

## 三、编码铁律（最重要，先读）

导出由加载项决定编码（UTF-8-BOM）；但**写回 Access 前，模块/类模块/宏必须转 GBK 无 BOM + CRLF，窗体用 UTF-16 LE + BOM**。编码错了报的是 SQL `3075` 这类业务错误，不是编码错误。完整规则见 `references/vba-encoding-rules.md`；.src 文件本身的编码/换行/新建规则见 `references/vcs-text-editing.md`。

## 四、加载项 API 的可操纵面

入口是 `Application.Run("<加载项全路径去掉扩展名>.API", 命令, 参数...)`，**只转发版本控制类的公共方法**（`Export`/`FullExport`/`ExportVBA`/`Build`/`MergeBuild`/`ExportObject`/`ImportObject`/`CompileVBA`/`IsVBACompiled`/`RunVBA`/`ExecuteSQL` 与 `GetVCSVersion` 等状态查询）；`modAPI` 里的其他内部函数不在转发范围，调了会报 VBA `438`。**常用命令表、Build/MergeBuild 的完整调用姿势与判据的展开单点在 `docs/troubleshooting.md` 第 20 节**，此处只留结论。

## 五、Build 的硬前提：宿主库必须与目标库同名同路径

`Build` 拿「当前打开的库全路径」与「源目录推导出的目标库全路径」（由 `vbe-project.json → Items.FileName` 决定）做大小写不敏感比较：一致才进入完全构建；不一致弹模态框阻塞调用；**当前没有库打开**则解析不了 `<全路径>.API` 同样失败。驱动脚本会在调用前自动校验，不同名直接拒绝。

## 六、MergeBuild 的前置条件

`MergeBuild` 有前置开关，不满足时弹框后**静默返回约 0.5s、无日志**，极易误判成功。**回写验证不要首选它，用 `ImportObject`**：对象级、同步、返回带 `logPath`、可当场复核。

## 七、Build / 回写还必须满足两个前提

1. **后端库（`Data.accdb` 或链接后端）必须在目标库同级目录**——缺了它链接表全建不出，且 Build 日志**不报"链接失败"**，只能靠事后漂移核验发现。
2. **调用后必须充分等待**——`Build` 外层 `Application.Run` 约 0.4 秒就返回，重活派发出去继续跑；用 `--settle` 等够，否则掐死没干完的活。

## 八、怎么判别"这次构建真的成功了"

「测试绿」不能单独证明构建成功。四判据须**同时满足**：① 出现 `Build_<时间戳>.log` 且含 `Created blank database for import`；② 出现被改名的备份库；③ 新库文件大小与原库**明显不同**；④ 对新库漂移核验为 0 项（含链接表）。

## 九、加载项崩溃判别法（三步）

调用非平凡 API 时若 Access 崩掉、报「远程过程调用失败 / RPC 服务器不可用」，先别改调用代码：① `tasklist` 看 `MSACCESS.EXE` 是否还在；② 事件日志 `Application Error` 看异常码（`0xc0000374` = 堆损坏）与 `%LOCALAPPDATA%\CrashDumps`；③ 阶梯探针（`scripts/probe_vcs_api.py`）逐个命令试，可见"前几个 OK、第几个一起就死"，并加对照复跑以排除守卫嫌疑。**根因多为加载项副本本身损坏**：简单命令（GetVCSVersion 等）正常、复杂命令崩，几乎都是副本问题，不是你的库或参数问题——一律用技能目录运行位 `Microsoft Access Version Control System\Version Control.accda`。完整步骤、判据与处置见 `docs/troubleshooting.md` 第 19 节。

## 十、MCP 注意事项

- 读写分离：只读 SELECT 用 `vcs_execute_sql`；调已有函数用 `vcs_call_vba`（默认开）；跑代理生成的代码用 `vcs_run_vba`（需开启 `McpAllowRunVBA`，见 1.2 第 6 步）。
- 编译失败时 MCP 报不出出错模块/行，需让用户在 Access 里「Debug → Compile」定位首错行再贴代码。
- 技能目录运行位以外的 add-in 路径、生产库建议设 `ACCESS_VCS_DISABLE_WRITES=true` 先只导出审查。
- 驱动脚本 `scripts/vcs_drive.py` 在主线程直接调加载项 API（绕开 MCP 的异步/线程坑），适合做"库到底能不能被完全操纵"的真机验证。

## 十一、NetUI 对话框键盘解锁（Build Name Conflict 等）

加载项弹的「Build Name Conflict」等对话框是 **Office Fluent/NetUI 框架**渲染：窗口类 **`NUIDialog`**、子控件类 **`NetUIHWND`**（文本为空）。按钮**不是标准 Win32 子窗口**——`EnumChildWindows` + `PostMessage(BM_CLICK)` **必然失效**，守卫的按钮匹配也找不到它，脚本会挂到超时。

**解法（实测有效）**：发键盘 `VK_RETURN` 触发默认「确定」按钮：

```python
hwnd = win32gui.FindWindow(None, "Build Name Conflict")
win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0x001C0001)
win32gui.PostMessage(hwnd, win32con.WM_KEYUP,   win32con.VK_RETURN, 0xC01C0001)
```

**教训**：遇到点不动的对话框先 `GetClassName`；只要看到 `NUIDialog`/`NetUIHWND`，立刻放弃 `BM_CLICK` 改键盘。这与"Build 宿主库与目标库同名同路径"（第五节）是配套的——名字不一致触发的正是这类 NetUI 模态框。

## 十二、CLI 与 MCP 的另外两个坑

1. **CLI 形态忽略 `output_dir`**：`msaccess-vcs` CLI 的某些命令会忽略传入的输出目录参数、写到默认位置——不要依赖 CLI 参数控制输出路径，脚本里显式用 `--db`/绝对路径并事后核对文件落点；不确定时走 `scripts/vcs_drive.py`（主线程直接调 API，参数全由脚本掌控）。
2. **MCP 会话会留 Access 窗口与锁**：MCP 驱动的 Access 实例用完可能不退出，残留 `MSACCESS.EXE` 与 `.laccdb` 会占用库、让后续刷新链接/构建假失败。每次 MCP 操作收尾：`taskkill /F /IM MSACCESS.EXE` + 删 `*.laccdb`（封装在 `scripts/fwguard.py::kill_access`），再核验进程=0、无锁文件。

## 十三、ExportObject 错误 91 修复记录（NoIndex 保护，勿回退）

**历史修复记录已归档 `changelog.md`（v2.12 条目）**，此处只留操作口径：`blnNoIndex:=True`（MCP/API 路径）下 `VCSIndex.Update` 在 `Me.Disabled` 时返回 `Nothing`，随后 `clsDbForm/clsDbModule/clsDbReport/clsDbVbeForm` 的 `Export` 无条件访问 `cIdx.FolderAnnotation` → VBA 错误 91「对象变量或 With 块变量未设置」，高亮行 `cIdx.FolderAnnotation = IIf(Len(strAnnotation) > 0, strAnnotation, FOLDER_ANNOTATION_NONE)`。源码保护（`If Not cIdx Is Nothing Then` 包裹）已合入 `msaccess-vcs-addin` 四处组件并构建进运行位 accda，**重构建时不得回退**；重构后必须用「导出单个对象 → 无错误 91、产物落盘、索引正确」复核（自动化门禁见 `scripts/rebuild_and_install.py`）。

## 十四、自动化测试：弹窗捕获/点击、报错与调试信息捕获、VBE 操纵

所有自动化一律代码化，禁止人工点弹窗。**四类手段（A 弹窗捕获/点击、B 代码报错捕获、C 调试信息捕获、D VBE 操纵）的执行细则以 `SKILL.md` 7.2 节为单点**（手段 A–E 逐条展开：`scripts/fwguard.py` 调用方式、`silent_execute`、`LogErr` 探针、`modDebugLog` 注入、`fwguard.get_proj` 定位、`probe_vcs_api.py` 阶梯探针），此处不再重复。前提：Access 选项开启「信任对 VBA 工程对象模型的访问」。
