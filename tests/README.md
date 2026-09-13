# tests/ 验证脚本

本目录放技能自身的通用验证脚本。目标是**零手动、可反复跑**，把"看起来对"变成"有断言证明对"。

## 当前脚本

### `test_skill_integrity.py`（纯 Python，无需 Access）
冒烟测试，校验：
1. SKILL.md 引用的 **12 个子项目目录**都存在；
2. `references/`、`templates/`、`scripts/`、`assets/`、`docs/` 关键文件齐全；
3. 模板编码符合 `references/iron-laws.md`：标准模块 UTF-8-BOM、类模块 GBK 无 BOM、宏无 BOM；
4. 技能自有文件**无自研框架专有残留**（clsLog / Framework_* / 门禁数字 / TS_* 链接表等；`changelog.md` 是唯一记载变化处，豁免扫描）。

运行：
```bat
python tests/test_skill_integrity.py
```
退出码 0 = 全绿；1 = 有失败项。

### `vcs_com_verify.py`（需要 Access + COM）—— VCS 通道回归（COM 层）
验证 VCS 加载项安装与运行位的完整可用性：
1. Ribbon COM 加载项已加载（`COMAddIns` 含 `MSAccessVCSLib.AddInRibbon` → Access 重启后出现 Version Control 选项卡）；
2. API 阶梯：`GetVCSVersion`=5.0.1、`IsDatabaseOpen`、`GetProjectName`；
3. `ExportObject module basSample` 返回 success（**错误 91 NoIndex 修复的回归门禁**）；
4. `CompileVBA`=True；全程无弹窗；收尾只清本次新拉起的 Access 实例（进程快照对比）。

```bat
python tests/vcs_com_verify.py --db <宿主库.accdb>
```
退出码 0 = 全绿；1 = 有失败项。

### `vcs_mcp_verify.py`（需要 Access + COM）—— VCS 通道回归（MCP 层）
启动 `msaccess-vcs-mcp` 服务器走真实 JSON-RPC，硬断言 6 步：
`initialize` → `tools/list`（21 工具）→ `vcs_list_objects` → `vcs_call_vba(Hello)=VCS OK` → `vcs_compile_vba` → `vcs_export_object`（success + logPath）。

```bat
python tests/vcs_mcp_verify.py --db <宿主库.accdb>
```
退出码 0 = 全绿；1 = 有失败项。

> **宿主库准备**：两个回归脚本都需要一个含标准模块 `basSample`（含 `Function Hello() As String: Hello = "VCS OK": End Function`）的 `.accdb`。
> 可复用 `examples/01_blank_db_to_vcs_loop` 的建库方式，或按 `templates/standard-module-stub.bas` 建模板库：
> `python scripts/rebuild_module_from_src.py <库.accdb> templates/standard-module-stub.bas`（导入后确认模块名与 Hello 函数）。

## 需要 Access + COM 的验证（在 `scripts/` 里）
这些验证要在本地装有 Microsoft Access 的环境跑，属于业务系统构建后的真机核验：
- `scripts/probe_templates.py`：建临时空库 → 导入模板 → 断言类型/乱码/行数 → 加载宏 → 删临时库。
- `scripts/vcs_consistency_check.py`：实时库 vs `.src` 漂移核验（退出码 0=无漂移、3=有漂移）。
- `scripts/diagnose_vba_project.py`：开工第一步跑工程能力矩阵，确认读/改/增删各能力是否可用。
- `scripts/probe_vcs_api.py`：VCS 加载项 API 阶梯探针，定位第几个命令导致 Access 崩溃。

> 自动化环境约定：所有无人值守脚本必须带弹窗守卫（`scripts/fwguard.py`），收尾清理 `MSACCESS.EXE` 进程与 `.laccdb` 锁文件。详见 `references/iron-laws.md`。
