# Access 自动化驱动与 VBE 工程操作（COM SOP）

> 本文是 `SKILL.md` 的**详尽配套**：SKILL.md 讲「先做什么后做什么」（主干），本文给具体怎么用 Python + win32com 把 Access 驱动起来、零手动跑完每一步的命令与标准动作，以及改 VBA 工程时每条真机实测的确定性规则。
> 硬规则以 `references/iron-laws.md` 为准；出错了按症状去 `docs/troubleshooting.md` 查。
> 版本控制（VCS 加载项 + msaccess-vcs-mcp）贯穿每个阶段，本文的命令都建立在「改 `.src` 文本 → 构建/回写回库 → 漂移核验」的闭环之上。

## 一、环境前置（一次性的）

- 本机已安装 Microsoft Access（2010 及以上，32/64 位均可）。
- 信任中心开启「信任对 VBA 工程对象模型的访问」（`AccessVBOM=1`），并把工作目录加进「受信任位置」。否则 Python 用 win32com 驱动时 `.VBProject` 报拒绝访问，AutoExec 宏与代码也不会自动跑。
- 就位两条自动化通道（贯穿后续每个阶段）：
  - **VCS 加载项运行位（技能目录自包含）**：确认 `Microsoft Access Version Control System\Version Control.accda` 在位，操纵一律优先用它（判别见 `references/vcs-three-channels.md` 的「加载项崩溃判别法」一节）。
  - **msaccess-vcs-mcp**：把加载项包装成 `vcs_export_database` / `vcs_import_object` / `vcs_rebuild_database` / `vcs_compile_vba` 等工具，供脚本与 AI 编码代理驱动；连接方式见 `msaccess-vcs-mcp/README.md`，项目根 `.env` 写 `ACCESS_VCS_DATABASE=<你的库.accdb>`。
- 所有无人值守脚本统一用 Python + win32com（**必须用 `EnsureDispatch` 才拿得到 VBE 对象**）驱动 Access，并统一套 `scripts/fwguard.py`，做到零手动、可重复。
- 解释器建议用明确路径的 Python 3.10+（如 `D:\Python\cpython-3.13-windows-x86_64-none\python.exe` 或 `uv run --python 3.13`）；`cscript/wscript` 与 PowerShell `New-Object -ComObject` 常被安全策略拦截，Python + win32com 是可用通道。

验证加载项已注册：

```powershell
Get-ItemProperty 'HKCU:\Software\VB and VBA Program Settings\MSAccessVCS\Add-in'
```

## 二、VBE 工程能力矩阵（实测）

| 操作 | 结果 | 前提 |
|---|---|---|
| 枚举组件 / 读 `CodeModule.Lines` | 可用 | 无 |
| `VBComponents.Import(.bas/.cls)` | 可用 | 文件必须 **GBK 无 BOM**；目标工程必须定位正确 |
| `VBComponents.Remove` | 可用 | 工程定位正确 |
| `VBComponents.Add(1/2)` | 可用 | 工程定位正确；但**新增后无法改名**（见铁律 4） |
| `LoadFromText(4, 宏名, 文件)` | 可用 | 文件必须 **GBK 无 BOM** |
| `ReplaceLine` / `InsertLines` / `DeleteLines` | 可用 | **必须先把该组件设为活动代码窗格**（见铁律 2） |
| `References.AddFromFile/AddFromGuid` | 不可靠 | 实测"假成功"（不抛异常但引用不进集合），勿依赖 |
| `Application.SaveAsText`（模块/宏） | 可用 | 这是 VCS 导出的底层通道，四类对象全部可用 |

## 三、COM 标准动作（SOP）

### 3.1 连接与定位工程

```python
import win32com.client, pythoncom, os
pythoncom.CoInitialize()
app = win32com.client.gencache.EnsureDispatch("Access.Application")  # 必须用 EnsureDispatch 才拿得到 VBE
app.Visible = True

def get_proj(app, db_path):
    for p in app.VBE.VBProjects:
        try: fn = p.FileName
        except Exception: fn = ""
        if fn and os.path.normcase(fn) == os.path.normcase(db_path):
            return p
    return None

app.OpenCurrentDatabase(r"<库.accdb>")
proj = get_proj(app, os.path.abspath(r"<库.accdb>"))   # 不要用 ActiveVBProject（被加载项时指向加载项）
```

### 3.2 原子重建模块（不手工粘贴代码）

```bash
# 不要 InsertLines 灌大段；用 Remove + Import（GBK 无 BOM）+ 读回比对
python scripts/rebuild_module_from_src.py "<库.accdb>" "<src 文件>"
```

脚本封装了完整五步：读出 `.src` 权威文本 → 转 GBK 无 BOM + CRLF 写临时文件 → 按 `Attribute VB_Name` 移除旧组件 → `Import` → 读回逐行比对。组件名取自文件头 `Attribute VB_Name`，**与文件名无关**；按文件名 Remove 会漏删、留僵尸模块。

### 3.3 漂移核验（每次改完 .src 必跑）

```bat
python scripts/vcs_consistency_check.py "<库.accdb>" "<库>.src"
```

退出码 0=无漂移、3=有漂移。逐类比对标准模块/类模块/窗体/宏/查询/表。判定构建是否真成功（不止看测试绿）见 `references/vcs-three-channels.md` 的「怎么判别构建成功」一节。

### 3.4 模板验证

```bat
python scripts/probe_templates.py
```

建临时空库 → 导入模板 → 断言类型/乱码/行数 → 加载宏 → 删临时库。

### 3.5 弹窗守卫与清理

所有无人值守脚本套 `scripts/fwguard.py`：按首字匹配关错误框、清理 `MSACCESS.EXE` 进程与 `.laccdb` 锁。`finally` 里 `CloseCurrentDatabase(); Quit()`。

**进程清理用「快照 + 只杀新进程」模式**（防止误杀用户正在使用的 Access，也避免留下高完整性孤儿）：

```python
import fwguard
before = fwguard.snapshot_access()      # 记录启动前的 Access PID
# ... 启动 COM 实例并执行操作 ...
app.Quit(2)
fwguard.kill_new_access(before, base_dir=base)   # 只终止本次新拉起的实例，返回清理不掉的孤儿 PID
```

`fwguard.kill_access(base)` 保留为向后兼容（无差别清理，仅用于受控脚本收尾；高完整性孤儿 `Access is denied` 杀不掉时不阻塞——它们不占 COM 单实例，重启系统后消失）。

坑点：错误框按钮文本常被渲染截断为单字（`继续`→`续`、`结束`→`束`），按全名匹配会失配，必须按**首字**匹配；个别错误框没有标准 Button 子控件，需退回 `WM_COMMAND + IDOK` / 回车。守卫还含"另存/保存/覆盖/替换/删除 字样一律点取消"的安全策略。**NetUI 系对话框（窗口类 `NUIDialog`/`NetUIHWND`）则完全点不动**——见 `references/vcs-three-channels.md` 的「NetUI 对话框键盘解锁」一节。

### 3.6 TDD 做法（逻辑层与界面层分离）

把可测逻辑写成不依赖界面的纯函数（标准模块），界面层只收输入、调内核、显示结果。对纯函数用 `Application.Run` 传参断言返回值；对整库用 `vcs_compile_vba` 保证可编译，再用漂移核验确认对象齐整。开工前用 `scripts/diagnose_vba_project.py` 探测工程能力矩阵。

### 3.7 自动化测试手段详单（弹窗/报错/调试/VBE 操纵，全代码化）

四类手段（A 弹窗捕获/点击、B 代码报错捕获、C 调试信息捕获、D VBE 操纵）的**定义与执行细则以 `SKILL.md` 7.2 节为单点**（手段 A–E 逐条展开，含 NetUI 解锁、`silent_execute`、`LogErr` 探针、`modDebugLog` 注入、`fwguard.get_proj` 定位），本文不再重复展开。此处仅保留本 SOP 语境下的调用入口：

- **弹窗捕获与点击**：`scripts/fwguard.py`（`start_guard()` 后台轮询 → `dialogs_seen()` 回读全部标题/正文供断言；按钮按**首字**匹配；NetUI 系对话框点不动，用键盘 `VK_RETURN`，见 `references/vcs-three-channels.md` 第十一节）。
- **代码报错捕获**：Python 侧 `silent_execute`（`except` 捕获 `com_error`，`e.args[2][2]` 取错误描述）；VBA 未处理错误框文本由 fwguard 捕获；定位出错行用 `On Error GoTo EH` + `LogErr`（`Err.Number/Description/Source/Erl`，模板 `templates/modDebugLog.bas`）。
- **调试信息捕获**：`modDebugLog.bas` 注入被测库（`rebuild_module_from_src.py` 原子导入）→ 被测代码 `LogPrint`/`LogTimer`/`Watch` → Python 读 `.\vba_debug.log` 断言。
- **操纵 VBE**：`fwguard.get_proj(app, path)` 按 `FileName` 定位 → `CodeModule` 读改（`ReplaceLine`/`InsertLines`/`DeleteLines`，须先激活代码窗格）→ 改完 `CompileVBA` → `ExportObject` 导出复核 → 漂移核验 0 项。

## 四、六条工程铁律（真机实测）

### 铁律 1：`ActiveVBProject` 不是你的工程

加载了 VCS 加载项时，`app.VBE.ActiveVBProject` 指向加载项工程（常密码保护），所有 `Add` / `Remove` / 写入都会打在加载项上，并抛出 **误导性错误**（"设备 I/O 错误"、"该工程已被保护"）。**必须**按 `FileName` 遍历定位（见 3.1 的 `get_proj`）。

### 铁律 2：原地写入必须先激活代码窗格

不激活就写，`InsertLines` / `AddFromString` / `ReplaceLine` 一律报 `STG_E_FILENOTFOUND`（表现为 `'文件未找到'`）。这是**最容易误判为"Access 环境不允许写入"**的坑。

```python
cm = comp.CodeModule
app.VBE.ActiveCodePane = cm.CodePane
comp.Activate()
cm.InsertLines(...)      # 现在可用
```

对照实验（同一库，五种写法，各自独立会话）：直接 `InsertLines` 失败；`ActiveCodePane` + `Activate` 后 `InsertLines` **成功**；`AddFromString` 失败；`ReplaceLine` 失败；只把 VBE 主窗置可见再写 失败。

### 铁律 3：不要用 `InsertLines` 灌大段代码，用原子替换

`InsertLines` 写大块文本会**部分写入**：注释头写进去了、正文没写进去，而且抛异常。这会在模块尾部留下**截断残片**，很难察觉。正确做法见 3.2 的 `rebuild_module_from_src.py`（Remove + Import + 读回比对，原子且可验证）。

`.cls` 读回比对时要先剥掉 `VERSION 1.0 CLASS` / `BEGIN` / `END` / `MultiUse` / `Attribute ...` 这些 VBE 不保存的头部行（约 9 行），否则会误报"不一致"。

### 铁律 4：`Add(2)` 能建类模块，但**改不了名**

实测 `VBComponents.Add(2)` 成功（type=2、默认名"类1"），但紧接着 `comp.Name = "clsX"` 抛 **VBA 错误 53（文件未找到）**。因此**不要走 `Add(2)+改名` 路线**。建类模块的可靠路线是 `Import` 一个 **GBK 无 BOM** 的 `.cls`（详见 `references/vba-encoding-rules.md`）。

### 铁律 5：Tab 控件只能 COM 构建

向已有 `.form` 文本插入 `Begin Tab` 块，`LoadFromText 2` 必失败——Access 解析器只接受它自己原子生成的 Tab 结构。

```python
tab = app.CreateControl(form_name, 123)          # acTabCtl=123，自动带 2 页
newp = tab.Pages.Add()                            # 无参
newp.Name = "页1"
ctl = app.CreateControl(form_name, ctype, pythoncom.Empty, newp.Name)  # Section 传 pythoncom.Empty，Parent 传页名字符串
```

### 铁律 6：门禁的"失败组"多数是环境态假象 + 脚本必须带弹窗守卫

- 同一份代码，测试可能报若干组失败，清掉残留 `MSACCESS.EXE` 与 `.laccdb` 后重跑就全绿。原因：残留进程占用后端库，刷新链接等操作失败。**跑测试前必做**：`taskkill /F /IM MSACCESS.EXE` + 删除 `*.laccdb`（封装在 `scripts/fwguard.py::kill_access`）。判定"真失败"与"假失败"：**同一会话内直接调该函数看返回值/错误号**，不要只看测试汇总。
- VBA 运行时错误会弹出模态框（标题 `Microsoft Visual Basic`，按钮 继续/结束/调试/帮助），无人值守脚本会被**永久阻塞**——必须套守卫（见 3.5）。

## 五、命令速查

| 动作 | 命令 |
|---|---|
| 导出全量源码 | `msaccess-vcs export <库> <出目录> --full`（或 `python scripts/vcs_drive.py FullExport <目录> --db <库>`） |
| 单对象导入回库 | `msaccess-vcs import <库> <目录> --object-types modules` 或 `vcs_import_object` |
| 从源码重建库 | `msaccess-vcs rebuild-database <目录> <新库>` |
| 主线程驱动 Build | `python scripts/vcs_drive.py Build <目录> --db <同名库> --settle 100` |
| 原子重建模块 | `python scripts/rebuild_module_from_src.py <库.accdb> <src 文件>` |
| 漂移核验 | `python scripts/vcs_consistency_check.py <库> <目录>` |
| 模板验证 | `python scripts/probe_templates.py` |
| 工程探测 | `python scripts/diagnose_vba_project.py <库.accdb>` |
| 回收站删除 | `python scripts/recycle.py <文件>`（`nul`/`con` 等设备名会让整批返回 124，须单独成批并先 `\\?\` 改名） |

> 命令中的 `msaccess-vcs` 是 MCP 暴露的 CLI 形态；等价地也可直接调 `scripts/vcs_drive.py` 或 MCP 工具（`vcs_export_database` / `vcs_import_object` / `vcs_rebuild_database` / `vcs_compile_vba`）。
