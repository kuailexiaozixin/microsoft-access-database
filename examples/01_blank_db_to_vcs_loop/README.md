# 示例 01：从空库到 VCS 闭环

**一句话**：新建一个空白 Access 库，往里导入一个模块，接着用版本控制把它导出成文本源 `.src`（单一真相源），最后用漂移核验证明「实时库 ↔ .src」完全对齐。这正好印证 `SKILL.md` 从阶段 1 起就接上源码管理的做法是成立的。

## 运行前提

- 本机已装 Microsoft Access（2010 及以上，32/64 位均可）。
- 已安装 Version Control 加载项（技能目录运行位）：检查 `Microsoft Access Version Control System\Version Control.accda` 是否存在。
- Python 已装 `pywin32`（`pip install pywin32`）。

## 运行

```bat
python run.py
```

## 它做了什么（全程代码，无手工点按）

1. **前提检查**：确认 `pywin32` 与加载项正式安装版在位。
2. **建空库**：用 `win32com` 的 `NewCurrentDatabase` 在临时目录建 `Sample.accdb`。
3. **写示例模块文本**：写一个自包含的 `basSample.bas`（UTF-8，带 `Attribute VB_Name`），里面只有一个 `Hello()` 函数。
4. **导入模块**：调 `scripts/rebuild_module_from_src.py`——它把文件转成 GBK 无 BOM + CRLF、按 `Attribute VB_Name` 原子替换旧组件、`Import` 进库、再读回逐行比对，绝不在 VBE 里手动粘贴。
5. **导出 `.src`**：调 `scripts/vcs_drive.py FullExport` 把库对象导出成文本源目录 `Sample.src`，作为单一真相源纳入版本管理。
6. **漂移核验**：调 `scripts/vcs_consistency_check.py` 比对实时库与 `.src`，退出码 0 = 无漂移。

## 你该看到的结果

- 第 4 步打印「导入 basSample … 一致」，证明原子导入成功、中文无乱码。
- 第 6 步打印「漂移合计 0 项 / 无漂移」，证明库与文本源对齐——这就是版本控制作为方法的闭环证据。

## 与技能文档的对应

- 为什么中文会乱码、`.src` 该用什么编码：读 `references/vba-encoding-rules.md`。
- 三套 VCS 工具怎么选、「怎么判别构建成功」：读 `references/vcs-three-channels.md`。
- 阶段 1「库一建好就导出 `.src`」：读 `SKILL.md`。
