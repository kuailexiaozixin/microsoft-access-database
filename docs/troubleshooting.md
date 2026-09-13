# Access 数据库开发排错手册（故障 → 根因 → 修复）

> 适用范围：基于本技能构建 Access + VBA 业务系统、用 VCS 通道操纵 `.accdb`、做自动化测试验证时遇到的典型错误。
> 所有诊断都用 Python 脚本（COM 驱动 Access + 弹窗守卫）完成，零手动操作。
> 编码、类模块、工程定位、弹窗守卫这几条硬规则先读 `references/iron-laws.md`；三套 VCS 工具的正确姿势读 `references/vcs-three-channels.md`。

---

## 1. 模块被静默丢弃：重开报「下标越界」/ `Application.Run` 报「找不到过程」

**现象**
- 关库重开后某模块不存在，`CurrentProject.AllModules` 索引不到。
- 所有 `Application.Run("Xxx")` 报「找不到过程」，即使代码在 `.src` 里明明存在。

**根因（高概率）**
- **条件编译引用了未定义常量**：`#If UNDEFINED_CONST = 1` 中常量未定义 → 编译错误 → 该模块在关库时被 Access 静默丢弃（不是报错，是直接没存进去）。
- 验证方法（最小复现）：空白库建一个含 `#If UndefinedConstX = 1` 的模块，插入 → 重开 → 模块不存在、`GetX` 报「找不到过程」。
- 等价根因还有：`.cls` 首行 `VERSION 1.0 CLASS` 头被 mojibake BOM（`ï»¿`）损坏 → VCS `Import` 建成 type=1（**BOM 才是真凶**，见第 13 节）→ 连带拖垮整库编译、全工程瘫痪。

**修复**
- 删除 `.src` 中指向未定义常量的 `#If/#ElseIf/#End If` 包裹，保留恒真的代码分支。
- 修复 BOM：剥除 `C3 AF C2 BB C2 BF`，写回真实 `EF BB BF`。
- 重建：**转 GBK 无 BOM + CRLF** 后 `Remove + Import`（`scripts/rebuild_module_from_src.py`），再重开验证「存在 = True」「Type=2」。
  （不要用 `Add(2) + InsertLines(1, code)`：**`InsertLines` 灌大段代码会部分写入并留残片**；`Add(2)` 也改不了名。）

**预防**
- 任何 `Application.Run` 报"找不到过程"时，先跑 VBE 全量编译定位编译错误，再查模块是否真存在（`CurrentProject.AllModules`）。

---

## 2. 错误 #3027 不能更新。数据库或对象为只读。

**现象**
- 窗体/记录集执行 `Update`/`AddNew`/`Delete` 时报 3027。

**三类根因（按优先级排查）**
1. **文件/链接层**：`.accdb` 文件属性勾了"只读"；或有 `.laccdb` 残留锁；或后端数据库只读；或链接表指向只读后端。
   - 排查：`attrib` 看只读位；`tasklist` 看有无孤儿 `MSACCESS`；删残留 `.laccdb`；`TableDefs` 查链接表 `Connect` 指向。
2. **逻辑非可更新记录集（最常见且最难查）**：RecordSource 是下列查询之一，Access 无法定位单行回写：
   - 含 `DISTINCT` / `GROUP BY` / `UNION` / 聚合函数的查询；
   - 多表 JOIN 且未包含每个基表的主键/唯一键；
   - `SELECT *` 跨多表 JOIN 时某些字段不可更新；
   - 基于上述查询再嵌套的查询。
   - 排查：枚举 `QueryDefs`，对每条 SQL 用关键字/`JOIN`/聚合做静态扫描，列出疑似非可更新查询，逐个在 Access 里打开看能否编辑。
3. **设计层**：窗体 `RecordsetType` 设为快照（Snapshot）而非动态集（Dynaset）；或控件绑定到表达式/计算字段。

**修复**
- 把只读查询改为可更新：补全各基表主键、拆分多表 JOIN 为子查询+绑定单表、或改用 `DAO` 直写基表而非绑查询。
- 窗体绑定改为单表/可更新查询；`RecordsetType = Dynaset`。

---

## 3. 类模块 type=2 丢失（被建成 type=1 标准模块）

**现象**
- `New clsX` 报"用户定义类型未定义"或"子程序或函数未定义"；`VBComponents` 里某类模块的 `Type=1`。

**根因**
- 真凶是**文件带 BOM**：带 `EF BB BF` 时首行 `VERSION 1.0 CLASS` 不被识别 → 退化成 type=1。
  把 `.cls` 写成 **GBK 无 BOM** 再 `Import`，实测得到 **type=2**。
- 另有独立根因：续行符 `_` 后插了空行 → 语法错误 + 级联"在 End 后只能出现注释" → 模块损坏。

**修复**
- 用 `scripts/rebuild_module_from_src.py`：读 `.src`（UTF-8-BOM）→ 转 **GBK 无 BOM + CRLF**
  → `Remove`（按 `Attribute VB_Name` 匹配，见第 17 节）→ `Import` → 读回逐行比对 → `DoCmd.Save 5`。
- ⚠️ **不要**用 `Add(2)` 建壳：能建 type=2 但**改不了名**（`comp.Name = "clsX"` 抛错误 53）。
- ⚠️ **不要**用 `InsertLines` 剥头灌正文：大段代码会"部分写入"并在尾部留残片。
- 修正续行符：`_` 后紧跟下一行，中间无空行。

---

## 4. BOM 乱码 `ï»¿`

**现象**
- 文本首行出现 `ï»¿VERSION 1.0 CLASS` 或 `锘緼ttribute VB_Name`。
- VCS 导入后模块/窗体损坏、工程编译失败。

**根因**
- 真实 UTF-8 BOM `EF BB BF` 被某环节按其他编码读后再以 UTF-8 存回，变成 `C3 AF C2 BB C2 BF`。

**修复**
- 以二进制读，剥除前 3 字节若为 `C3 AF C2 BB C2 BF`，改写 `EF BB BF`（或整文件按 utf-8-sig 重存）。
- 导入前按对象类型转码（见 `references/vba-encoding-rules.md` 与 `references/iron-laws.md` 的 A 节）。

---

## 5. 「该工程已被保护，不能执行对象」

**现象**
- 遍历 `VBProjects` 操作某库时被拒。

**根因**
- `VBE.ActiveVBProject` 指向了已加载的加载项工程（常密码保护），并非目标库。

**修复**
- 遍历 `VBProjects`，按 `FileName` 匹配目标 `.accdb` 再操作，不要依赖 `ActiveVBProject`。详见 `references/vba-com-automation.md` 的「`ActiveVBProject` 不是你的工程」一节。

---

## 6. Tab 控件插入失败（`-2146826003` / `-2146825338`）

**现象**
- 向已有 `.form` 文本插入 `Begin Tab` 块，`LoadFromText 2` 报错。

**根因**
- Access 解析器只接受它自己原子生成的 Tab 结构，文本方式插入必失败（与重命名/重缩进/版本号无关）。

**修复**
- 用 COM 构建：`app.CreateControl(name, 123)`（acTabCtl，自动带 2 页）→ `tab.Pages.Add()`（无参）→ `newp.Name=` → `app.CreateControl(name, ctype, pythoncom.Empty, page.Name)`（Section 传 `pythoncom.Empty`，Parent 传页名字符串）。完整做法见 `references/vba-com-automation.md` 的「Tab 控件」一节。

---

## 7. 自动化脚本挂起 / 弹窗堆积

**现象**
- 脚本长时间无输出，MSACCESS 进程不退出。

**根因**
- 编译错误/另存为对话框是模态的，无人点击会永久阻塞。

**修复**
- 弹窗守卫只关 `#32770` 普通对话框，绝不关 VBE 主窗（`wndclass_desked_gsk`）；含"另存/保存/覆盖/替换/删除"字样的对话框一律点"取消"，避免误写文件。
- `finally` 中 `CloseCurrentDatabase → Quit(2) → taskkill /F /IM MSACCESS.EXE → 删 .laccdb`。
- **按钮文本必须归一化**：Access 按钮常写成 `确定(&O)`，精确匹配 `确定` 会失配 → 守卫点不到 → `app.Run` 永久挂起。
  用 `scripts/fwguard.py` 的 `start_guard()`，它去掉 `(&X)`/`(X)`/`&` 后**按首字**匹配。

---

## 8. 结构性变更全部被阻塞：「设备 I/O 错误」（-2146828231 / 1000057）

**现象**
- `VBComponents.Add(1)`、`Add(2)`、`Remove`、`Import` 全部报 `设备 I/O 错误`。
- `References.AddFromFile / AddFromGuid` **不抛异常，但引用不会真的进入集合**（假成功）。
- 枚举组件、读代码、`ReplaceLine` 改已有代码却都正常。

**根因**
- 主库 VBA 工程被误当成「拒绝结构性变更」——实际上多数是**定位错工程**导致的假象：
  `ActiveVBProject` 指向的是 VCS 加载项（常加密），增删改全打在加载项上，目标库毫无变化，却报出 `设备 I/O 错误` / `该工程已被保护` 等误导信息。
- 已排除的其它可能：文件锁与孤儿进程、`/decompile` 反编译、先触发编译再 Add（编译无报错，Add 仍失败）。

**正确做法**
1. **先按 `FileName` 遍历 `VBProjects` 定位目标库**，不要用 `ActiveVBProject`（见 `references/vba-com-automation.md`）。
2. 改动后**必须重开复验**，不能只信会话内返回值（同会话里对象在、关库就没了的情况真实发生过）。
3. 重建模块用 `Remove + Import`（GBK 无 BOM），不要再用 `Add(2) + InsertLines`。

**开工第一步**：跑 `scripts/diagnose_vba_project.py` 做能力矩阵，确认读/改/增删各能力是否真的可用，避免死路试错。

---

## 9. 假线索：`CreateObject("ADODB.Stream")` 不等于缺 ADODB 引用

**现象**
- 静态扫描发现代码用了 `ADODB.Stream` / `WScript.Shell` / `Shell.Explorer`，于是去补 ADODB / WSH / Internet Controls 引用，但工程依然不编译。

**根因**
- 这些是**字符串后期绑定**（`CreateObject("...")` / `Array("Shell.Explorer.2", ...)`），编译期不需要任何引用，只有 `Dim x As ADODB.Stream` 这类**声明**才需要。

**判定规则**
- `Dim x As ADODB.Stream` → 需要引用（编译期解析类型）。
- `CreateObject("ADODB.Stream")` → 不需要引用（运行期按名创建）。

**教训**
- 看到库前缀先分清是「类型声明」还是「字符串后期绑定」，别急着补引用。

---

## 10. `InsertLines` / `ReplaceLine` / `AddFromString` 报「文件未找到」

错误码 `-2146828235` / `-2147287038`（`STG_E_FILENOTFOUND`），提示信息里还带一个不存在的 VBA 帮助文件路径 —— 属于**误导信息**。

**真因**：没有把目标组件设为活动代码窗格。

**修复**：

```python
cm = comp.CodeModule
app.VBE.ActiveCodePane = cm.CodePane
comp.Activate()
cm.InsertLines(...)          # 现在可用
```

> 更稳妥的做法是不要用 `InsertLines` 灌大段代码（会部分写入留残片），改用 `Remove + Import` 原子替换。详见 `references/vba-com-automation.md` 的「原地写入」与「原子替换」两节。

---

## 11. 导入后中文变乱码，运行时 SQL 报 `3075 语法错误(操作符丢失)`

**真因**：导入文件写成了 UTF-8（无 BOM），Access 按系统 ANSI（GBK）读，中文被解码错。

典型报错正文：

```
运行时错误 '3075': 语法错误 (操作符丢失) 在查询表达式 ''鐢?,'TDD鎺㈤拡','system'...' 中。
```

**修复**：导入文件必须 **GBK 无 BOM + CRLF**。导入后**读回逐行比对**验证。完整编码铁律见 `references/vba-encoding-rules.md`。

---

## 12. `.cls` 导入后变成标准模块（type=1）

**真因**：`.cls` 文件带 **BOM**，首行 `VERSION 1.0 CLASS` 不被识别。

**修复**：去掉 BOM、用 GBK 编码再 `Import` → 实测得到 type=2。

**易被误导的情形**：某个 `.cls` 用 UTF-8 无 BOM 导入也成功了且 type=2 —— 那是因为它**整份没有非 ASCII 字符**。一旦类里有中文就会乱码。不要据此认为编码无所谓。

---

## 13. 测试/门禁出现失败组

**先怀疑环境态**：残留 `MSACCESS.EXE` 或 `.laccdb` 占用后端库。

**处置顺序**：
1. `taskkill /F /IM MSACCESS.EXE` + 删除 `*.laccdb`；
2. 重跑测试/门禁；
3. 仍失败才按真缺陷排查，并且要**在同会话内直接调用该函数**看返回值/错误号，不要只依赖汇总（汇总只给计数，不给明细）。

---

## 14. 脚本无故卡死不返回

**真因**：VBA 运行时错误弹出了**模态错误框**，无人点击。

- 按钮通常是 继续/结束/调试/帮助，但**文本常被截断为单字**（`续`/`束`/`试`/`助`），按全名匹配会失配 → 必须按**首字**匹配；
- 个别错误框没有标准 Button 子控件 → 退回 `WM_COMMAND + IDOK` 与回车键；
- 统一用 `scripts/fwguard.py` 的 `start_guard()`。

---

## 15. 晚期绑定调用 `ParamArray` 方法只传 1 个参数 → 错误 13

见 `references/vba-com-automation.md` 的对应节。要点：传 ≥2 个参数即可；或另加 `ByVal` 单参包装方法。

---

## 16. 重建模块后库里有「僵尸模块」/ 改名没生效

**现象**
- 用脚本重建模块后，库里出现两个功能相同的模块（一个旧名一个 Attribute 里的名），或"明明改了文件名和代码，组件名却没变"。

**真因**
- `VBComponents.Import` 建出的组件名取自**文件头部的 `Attribute VB_Name`**，**不是文件名**。
  - 模板 `class-module-stub.cls`（头部 `Attribute VB_Name = "clsTemplate"`）导入后 → 组件名 `clsTemplate`。
  - 模板 `standard-module-stub.bas`（头部 `"basTemplate"`）导入后 → 组件名 `basTemplate`。
- 所以"按文件名移除旧组件"在两者不一致时会**漏删**，留下僵尸模块。

**修复**
- 移除旧组件、`DoCmd.Save 5` 一律**按 `Attribute VB_Name` 匹配**。
  `scripts/rebuild_module_from_src.py` 已按此加固：先解析头部拿组件名，导入后再核对 `c.Name` 是否与预期一致，不一致立刻打警告。

---

## 17. 实时库与 `.src` 悄悄脱节（源码里有的，库里没有）

**现象**
- `.src` 里写了某行调用或某过程，但实时库的对应模块是更旧的版本、根本不含这行。按 `.src` 判断"库里有什么"会得出错误结论。

**真因**
- 实时库是用**更早的 `.src`** 导入的（stale）：`.src` 更新了，但没人重建主库。

**修复 / 预防（常规化核验）**

```bat
python scripts/vcs_consistency_check.py "你的库.accdb" "你的库.accdb.src"
```

逐类比对 标准模块 / 类模块 / 窗体 / 宏 / 查询 / 表，任何"只在实时库"或"只在 `.src`"都会列出来，退出码 0=无漂移、3=有漂移。**每次改完 `.src` 后都该跑一次**。

---

## 18. 无关弹窗干扰（如「另存为」「DAL=on」、下载器窗口）

**现象**
- 守卫日志里出现与本工程无关的弹窗，例如 `'另存为' 正文=DAL=on` 或某个下载器窗口。

**判定**
- 这类窗口不是 Access 的：正文与按钮都不是 Access 的话术。守卫按「另存/保存/覆盖/替换/删除」危险字样策略点了**取消**，未造成任何写入；验证结果不受影响。

**处置**
- 不要因为看到陌生弹窗就怀疑脚本出错。判断标准是：**断言是否全绿 + 收尾是否干净（`MSACCESS.EXE` 归零、无 `.laccdb`）**。
- 若某弹窗反复出现并伴随失败，再单独抓窗口类名与标题定位来源。

---

## 19. 调用 VCS 加载项 API 时 Access 崩掉：「远程过程调用失败」/「RPC 服务器不可用」

**现象**
- `app.Run(r"...\Version Control.API", "Build", ...)` 报
  `(-2147023170, '远程过程调用失败。')`，下一次调用变成
  `(-2147023174, 'RPC 服务器不可用。')` —— COM 对象彻底没了。
- 表面看像"参数不对"或"库有问题"，很容易误判成"Access 库不可操纵"。
- 简单命令却完全正常：`GetVCSVersion`、`IsDatabaseOpen`、`GetProjectName` 都返回正确值。

**根因（实测定位）**
- **是加载项那份 `.accda` 副本本身坏了**，不是调用姿势、不是库、不是弹窗守卫。
- 判据一（事件日志）：`MSACCESS.EXE` 以异常码 **`0xc0000374` = 堆损坏（HEAP_CORRUPTION）** 崩在 `ntdll.dll`。这是原生层内存越界，纯托管调用不可能造成。
- 判据二（崩溃转储）：`%LOCALAPPDATA%\CrashDumps\MSACCESS.EXE.*.dmp` 与运行时刻一一对应。
- 判据三（阶梯探针）：`scripts/probe_vcs_api.py` `<库>` 逐个命令试，可见"前几个 OK，第几个一起就死"，且**关掉弹窗守卫（`--noguard`）结果完全相同** → 排除守卫嫌疑。

**处置**
- **一律使用技能目录运行位**：`Microsoft Access Version Control System\Version Control.accda`（安装位=运行位，自包含）。
  - `scripts/vcs_drive.py` 默认用它（运行位即技能目录内；`--skill-addin` 参数已不再必要，保留仅为兼容）。
- 排查手法固定成三步：`tasklist` 看进程是否还在 → 事件日志看 `0xc0000374` → 阶梯探针定位第几个命令死。
- 不要因为「简单命令通、复杂命令挂」就去改调用代码；先去比对加载项副本。更完整的判别法见 `references/vcs-three-channels.md` 的「加载项崩溃判别法」一节。

---

## 20. VCS 加载项：构建 / 回写的正确调用姿势（实测）

**可操纵面**
- `Application.Run("<加载项全路径去掉扩展名>.API", 命令, 参数...)` 只会转发到**版本控制类的公共方法**；加载项内部其它函数（如 `Preload`、`SetInteractionMode`）**不在转发范围内**，调了会报 VBA `438 对象不支持该属性或方法`。
- 常用命令与签名：
  | 命令 | 说明 |
  |---|---|
  | `Export` / `FullExport` / `ExportVBA` | 导出源码（`Export` 只导变更，`FullExport` 全量） |
  | `Build(strSourceFolder)` | **完全构建**：备份原库 → 建空白库 → 导入全部源码 |
  | `MergeBuild()` | 增量合并（**有前置条件**，见下） |
  | `ExportObject(类型, 名)` / `ImportObject(类型, 名)` | 单个对象导出 / 回写 |
  | `CompileVBA()` / `IsVBACompiled()` | 编译 / 查询编译状态 |
  | `RunVBA(代码)` / `ExecuteSQL(sql, maxRows)` | 执行 VBA / SQL |
  | `GetOption(名)` / `SetOption(名, 值)` | 读写选项 |
  | `GetVCSVersion` / `GetExportFolder` / `GetProjectName` / `IsDatabaseOpen` / `GetLogContent` | 状态查询 |

**Build 的硬前提：宿主库必须与目标库同名同路径**
- 构建引擎会拿「当前打开的库全路径」和「源目录推导出的目标库全路径」做大小写不敏感比较：
  - 一致 → 进入完全构建（它自己会关库、重建、再导入）；
  - 不一致 → **模态询问**，阻塞调用（表现为 RPC 中断或"什么都没发生"）；
  - **当前没有库打开** → Access 解析不了 `<全路径>.API` 库引用，同样失败。
- 目标库路径由源目录里的 `vbe-project.json → Items.FileName` 决定，最终 = `FSO.GetAbsolutePathName(源目录\..\FileName)`。
- `scripts/vcs_drive.py` 会在调用前自动校验，不同名直接拒绝执行。

**MergeBuild 的前置条件**
- 检查主界面 `chkFullBuild.Enabled`；为假时弹「必须先做一次完全构建或完全导出」，然后就什么都不做（约 0.5 秒静默返回）。
- 所以**回写验证不要首选 `MergeBuild`**，用 `ImportObject`：对象级、同步、返回带 `logPath`、可当场复核。

**回写端到端验证范式**
1. 在源码里给某个已有模块追加一个可识别的探针函数（保持 UTF-8-BOM + CRLF，只写 ASCII）；
2. `vcs_drive.py ImportObject modules <模块名> --db <库> --then-run <探针函数名>`；
3. 期望输出含 `logPath` 且复核函数返回预期值。
- 对应日志落在**目标库的** `ExportFolder\logs\Merge_<时间戳>.log`。

**`Build` 还有两个必须满足的前提（否则会"看起来成功、其实不完整"）**

1. **后端数据库（如 `Data.accdb`）必须在目标库的同级目录**。源里链接表按目标库所在目录解析。少了它，链接表全部建不出来，库不完整，而且 Build 日志里**不会显式报"链接失败"**，只能事后用漂移核验发现。
2. **调用后必须充分等待**。`Build` 的外层 `Application.Run` **约 0.4 秒就返回**（重活派发出去继续跑），完整构建实际要更久。所以要等够（如 `--settle`），否则可能掐死没干完的活。

**怎么判定"这次构建真的成功了"（四条缺一不可）**

| 判据 | 说明 |
|---|---|
| 出现 `Build_<时间戳>.log` | 且含 `Created blank database for import`、末尾有 `TOTAL RUNTIME` |
| 出现被改名的备份库 | 原库被改名留档 |
| 新库文件大小与原库明显不同 | 说明确实经过了"清空→重建→导入" |
| 对新库跑漂移核验 = 0 项 | 含全部表与链接表 |

⚠️ **"测试绿"不能单独证明构建成功**——若宿主库本身已是可用副本，构建其实没跑也照样绿。必须同时看上面四条判据。完整版见 `references/vcs-three-channels.md` 的「怎么判别构建成功」一节。

---

## 21. 窗体代码写回后关库回滚

**现象**
- 用 VBE 或脚本改了窗体类模块代码，关库重开后改动没了。

**根因**
- 改完窗体代码后没紧跟 `DoCmd.Save acForm, 窗体名`，设计缓存会在关库时回滚。

**修复**
- 任何写回窗体代码的动作之后，立即执行 `DoCmd.Save acForm, 窗体名`（COM 或 VBA 均可）。详见 `references/iron-laws.md` 与 `references/vba-com-automation.md`。
