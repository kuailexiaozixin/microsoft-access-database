# Access VBA / VCS / 自动化 铁律汇总

> 这些是跨多轮真机验证、反复踩坑后沉淀的确定性规则。任何构建 / 排错都先对照本文件。
> 配套：编码细节见 `references/vba-encoding-rules.md`；用 Python/COM 操纵工程的实测与命令见 `references/vba-com-automation.md`；三套 VCS 工具怎么选、构建怎么判成功、NetUI 对话框怎么解锁见 `references/vcs-three-channels.md`。

## A. 编码与导入（完整规则、危害与自检代码见 `references/vba-encoding-rules.md`，此处只留结论）

1. `.src` 源码一律存 **UTF-8-BOM**（`EF BB BF`）。
2. 模块（`.bas`）`LoadFromText 5` 或 `VBComponents.Import` 前必须转 **GBK 无 BOM + CRLF**。带 BOM 直导 → 首行 `锘緼ttribute` → 工程瘫痪、所有 `Run` 报"找不到过程"；写成 UTF-8 无 BOM 导入 → 中文全乱码 `鐢?`，含中文的 SQL 字符串被破坏 → 运行时报 `3075 语法错误(操作符丢失)`。
3. 窗体（`.form`）`LoadFromText 2` 必须 **UTF-16 LE + CRLF + BOM**。
4. 导入用 `LoadFromText` 覆盖式导入（不要先删空模块，空 VBA 工程会不可用）；导入前**必须关闭已打开窗体**（否则静默卡死）。
5. 窗体级 `On*` 事件属性必须在 `Begin Form` 之后、默认属性 `Begin` 块之前，否则报 `-2146826003 不能创建输出文件`。
6. `Version =21` 的窗体需归一化为 `Version =20` 才能被当前 Access 解析。

## B. 类模块与对象模型

7. 类模块必须 `vbext_ct_ClassModule`（type=2）。**真凶是文件带 BOM**：带 `EF BB BF` 时首行 `VERSION 1.0 CLASS` 不被识别 → 退化成 type=1。把 `.cls` 写成 **GBK 无 BOM** 再 `Import`，实测得到 **type=2**。
   另实测：`Add(2)` 能建 type=2，但随后 `comp.Name = "clsX"` 抛 **VBA 错误 53（文件未找到）**，**改名不可用** → 不要走 `Add(2)+改名` 路线。
8. 不要用手工剥 `.cls` 头再 `InsertLines(1, code)` 来建类模块（大段代码会**部分写入**留残片）。改用 `Import`（GBK 无 BOM）+ 读回逐行比对，见 `scripts/rebuild_module_from_src.py`。
   读回比对时 `.cls` 要剥掉 VBE 不保存的头部行，否则误报"不一致"。
9. 续行符 `_` 后必须紧跟下一行，中间插空行 → 语法错误 + 级联"在 End 后只能出现注释"。
10. 窗体类模块 `Type=100`（不是 `vbext_ct_MSForm=3`）；`Application.Modules` 只含已打开模块，测"模块存在"必须用 `CurrentProject.AllModules`。

## C. 编译与调用

11. 一个编译错误 = 整个 VBA 工程瘫痪，所有 `Run` 报"找不到过程"。遇"找不到过程"先查编译错误，再查模块是否真存在。
12. 对外接口统一用 **Function**（不要用 Sub）。`Application.Run` 调 Sub 时 `result = Run(...)` 必报错并触发重试，导致 Sub 被执行两遍（进度/计数类危险）。
13. 值可能是对象 → 用 `Sub + ByRef out`，按 `IsObject` 选 `Set`/等号，禁等号接函数返回值（错 450）。

## D. 取整与分层

14. 向下取整用 `Int()`，禁 `CLng/CInt`（银行家舍入，边界数据进错箱且不报错）。
15. 逻辑层与界面层分离：先写纯计算内核标准模块（无界面依赖，可自动化覆盖），界面层复用同一内核。

## E. 窗体与 Tab 控件

16. 单一真相源是 `.src`；窗体文本只保留默认样式块；`TabIndex` 0 连续。
17. Tab 块只能 COM 构建（`CreateControl(name,123)` → `Pages.Add()` → `CreateControl(name,ctype,pythoncom.Empty,page.Name)`）；文本插 `Begin Tab` 必失败。见 `references/vba-com-automation.md` 铁律 5。
18. VBE 写回窗体代码后必须紧跟 `DoCmd.Save acForm, 窗体名`，否则设计缓存关库回滚。

## F. 自动化环境

19. COM 必须 `win32.gencache.EnsureDispatch("Access.Application")` 才拿得到 VBE（晚绑定 `VBE=None`）。
20. 解释器建议用明确路径的 Python 3.10+；`python` 可能不在 PATH 时用 `uv run --python 3.13` 或直接指定解释器。
21. `cscript/wscript` 与 PowerShell `New-Object -ComObject` 常被安全策略拦截；Python + win32com 是驱动 Access 的可用通道。
22. 脚本 `finally: CloseCurrentDatabase(); Quit()`；收尾 `taskkill /F /IM MSACCESS.EXE` + 清残留 `.laccdb`；交付前确认 MSACCESS 进程=0、无 `.laccdb`、无挂起弹窗。
23. VCS 加载项：`Application.Run("<路径>.accda!DummyFunction")` 加载；API 入口是 `<加载项全路径去掉扩展名>.API`；`Build` 是异步的，调用后必须充分等待（见 `references/vcs-three-channels.md`）。

## G. 复刻第三方框架

24. `VBE.ActiveVBProject` 在被加载项时指向加载项（常密码保护）→ 误报"工程被保护"；遍历 `VBProjects` 按 `FileName` 匹配目标库。
25. 命名差异 ≠ 功能缺失（先大小写/前缀归一化再下结论）；统计真实对象须剔除 `~sq_` 查询与 `MSys` 表。
26. 四类资产四做法：自包含原样复刻；仅错误处理依赖核心 → 机制复刻+适配；代码依赖编译库 → 只复刻设计剥离代码；编译库无源码 → 按行为重写（不冒充复刻）。
27. 跨库复刻表用 DAO 原样重建（读 `Field.Name/Type/Size/Attributes`，文本带 Size，自动编号 `Attributes=17`），逐字段+逐索引+读写冒烟验证。

## H. 工程写入与弹窗守卫

28. 原地写入（`InsertLines`/`ReplaceLine`/`AddFromString`）**必须先激活代码窗格**（`ActiveCodePane` + `Activate`），否则报 `STG_E_FILENOTFOUND`（`'文件未找到'`），极易误判成"环境不允许写入"。详见 `references/vba-com-automation.md` 铁律 2。
29. **禁用 `InsertLines` 灌大段代码**：会**部分写入**（注释头进去了、正文没进去）并抛异常，在模块尾部留下**截断残片**。一律 `Remove + Import`（GBK 无 BOM）原子替换 + 读回比对。详见 `references/vba-com-automation.md` 铁律 3。
30. **脚本必须带弹窗守卫**（`scripts/fwguard.py`）：运行时错误框会永久阻塞无人值守脚本；按钮文本常截断为单字（`继续`→`续`），须按**首字**匹配；"另存/保存/覆盖/替换/删除"字样一律点取消。详见 `references/vba-com-automation.md` 3.5。
31. 门禁"失败组"多为环境态假象（残留 `MSACCESS.EXE` / `.laccdb` 占用后端库）；判定真假失败看**同会话内直接调该函数的返回值**，别只看汇总。详见 `references/vba-com-automation.md` 铁律 6。

## I. 组件命名、漂移核验与模板验证

32. **`VBComponents.Import` 的组件名取自文件头的 `Attribute VB_Name`，不是文件名**；移除旧组件、`DoCmd.Save 5` 一律按 Attribute 名匹配，否则漏删、留僵尸模块（`scripts/rebuild_module_from_src.py` 已加固）。详见 `references/vba-com-automation.md` 3.2。
33. **实时库与 `.src` 必须做漂移核验**，别凭记忆断言"库里有什么"：`scripts/vcs_consistency_check.py <库> [.src]`，退出码 0=无漂移、3=有漂移，每次改完 `.src` 都该跑。命令见 `references/vba-com-automation.md` 3.3。
34. **模板必须真机验证，不能"看起来对"就交付**：`scripts/probe_templates.py` 建临时空库 → 导入模板 → 断言类型/乱码/行数 → 删临时库。见 `references/vba-com-automation.md` 3.4。

## J. VCS 加载项操纵

35. **只用技能目录运行位的加载项**：`Microsoft Access Version Control System\Version Control.accda`（安装位=运行位，自包含）。其他目录里的副本可能与运行位不一致，是非平凡 API 崩溃（`0xc0000374` 堆损坏）的首要根因。
36. **`Build` 要求宿主库与目标库同名同路径**：目标库由源目录 `vbe-project.json → Items.FileName` 决定，与源目录名无关；不一致 → 加载项弹模态框卡死调用；没库打开 → 解析不了 `<全路径>.API` 库引用。
37. **回写验证首选 `ImportObject`，不要用 `MergeBuild`**（后者有前置条件，不满足时弹框后**静默返回约 0.5s、无日志**，极易误判成功）。
38. **构建/回写还有两个前提**：① 后端库（`Data.accdb` 或链接后端）必须在目标库同级目录，否则链接表全建不出且 Build 日志不报错，只能靠事后漂移核验发现；② 调用后必须**充分等待**——`Build` 外层 `Application.Run` 约 0.4 秒就返回，完整构建实际要更久，用 `--settle` 等够。
39. **"测试绿"不能单独证明构建成功**（宿主库若本身就是可用副本，构建没跑测试也照过）。四判据：有 `Build_<时间戳>.log` 且含 `Created blank database for import`；有被改名的备份库；新库大小与原库明显不同；对新库漂移核验 0 项。
40. **`API` 只转发版本控制类的公共方法**；别的内部函数不在转发范围，调了会报 VBA `438`。选命令前先核对可操纵面。
41. **NetUI 对话框（窗口类 `NUIDialog`/`NetUIHWND`）按钮不是标准 Win32 子窗口**，`EnumChildWindows`+`BM_CLICK` 必然失效——用键盘 `VK_RETURN` 触发默认按钮；MCP/CLI 还有 output_dir 与残留窗口的坑。展开见 `references/vcs-three-channels.md` 第十一节。
42. **脚本路径一律相对技能根计算，禁用环境路径**：所有 Python 脚本用 `os.path.dirname(os.path.dirname(os.path.abspath(__file__)))` 得技能根，再拼子路径；**禁止** `os.environ["APPDATA"]`、`os.path.expanduser("~")` 等指向技能目录之外的默认路径（旧 `%AppData%\MSAccessVCS` 正式位已归档清理，引用它会断链；历史实例教训见 `references/vba-com-automation.md` 3.5）。
43. **进程清理用「快照 + 只杀新进程」模式**：`before = fwguard.snapshot_access()` → 操作 → `fwguard.kill_new_access(before)`；无差别 `taskkill /F /IM MSACCESS.EXE`（`kill_access`）会误杀用户正在使用的 Access，且清理不掉高完整性孤儿（`Access is denied`，不占 COM 单实例，重启系统后消失，勿反复重试）。
