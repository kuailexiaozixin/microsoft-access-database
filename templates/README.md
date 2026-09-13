# templates/ 模板索引

这些是"抄了就能用"的确定性存根，全部按 `.src` 存储编码落盘。

| 模板 | 用途 | 存储编码 | 写进 Access 前的处理 |
|---|---|---|---|
| `class-module-stub.cls` | 新建类模块（type=2） | UTF-8-BOM | 转 **GBK 无 BOM + CRLF** 后 `Remove` + `Import` |
| `standard-module-stub.bas` | 新建标准模块（type=1） | UTF-8-BOM | 转 **GBK 无 BOM + CRLF** 后 `LoadFromText 5` / `Import` |
| `autoexec-macro.macro` | 打开即自动执行入口宏 | 无 BOM | 原样 `Application.LoadFromText 4, "AutoExec", 文件` |
| `src-structure.md` | `.src` 目录结构与编码对照 | Markdown | 只读参考 |

一键重建（自动完成编码转换 + 读回比对）：

```bat
python scripts/rebuild_module_from_src.py "<主库.accdb>" "<要重建的 src 文件>"
```

> 铁律：`.src` 存 UTF-8-BOM；写进 Access 一律 GBK 无 BOM + CRLF（窗体除外，用 UTF-16 LE+BOM）。
> 编码错了报的是业务错误（中文乱码 → SQL 3075），不是编码错误。
>
> 组件名取自模板文件头部的 `Attribute VB_Name`（不是文件名）：`class-module-stub.cls` → `clsTemplate`，
> `standard-module-stub.bas` → `basTemplate`。要改名字改那一行。
>
> 模板可用性由 `scripts/probe_templates.py` 真机验证（临时空库导入 + 类型断言 + 中文无乱码 + 宏可加载），
> 当前 8 项断言全部通过。
