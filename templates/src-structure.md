# .src 目录结构（单一真相源）

`YourDb.accdb.src/` 是 VCS 导出后的源码目录，也是重建 Access 库的唯一真相源。
**判断"库里到底有什么"永远以实时库 + 本目录为准，不要凭记忆下结论。**

```
YourDb.accdb.src/
├── AGENTS.md                 # VCS 工程元数据（工具识别用）
├── dbs-properties.json       # 数据库属性（含条件编译常量等）
├── documents.json            # 文档对象清单
├── project.json / vbe-project.json / vbe-references.json / vcs-options.json
├── vcs-index.idx             # VCS 索引（导出/比对用）
├── modules/                  # VBA 模块源码
│   ├── *.bas                 # 标准模块（type=1）
│   └── *.cls                 # 类模块（type=2，首行 VERSION 1.0 CLASS）
├── forms/                    # 窗体文本
│   ├── *.form                # 窗体定义（LoadFromText 2 导入）
│   └── *.cls                 # 窗体代码模块（Type=100）
├── macros/                   # 宏（LoadFromText 4 导入）
│   └── *.macro
├── queries/                  # 查询（*.sql + *.json 成对）
├── tbldefs/                  # 表定义（DAO 重建依据）
└── themes/                   # Office 主题
```

## 实测编码约定

| 对象 | `.src` 存储编码 | 写进 Access 时的编码 | 导入通道 |
|---|---|---|---|
| 标准模块 `.bas` | UTF-8-BOM | **GBK 无 BOM + CRLF** | `LoadFromText 5` 或 `Import` |
| 类模块 `.cls` | UTF-8-BOM | **GBK 无 BOM + CRLF** | `Remove` + `Import`（得到 type=2） |
| 窗体 `.form` | UTF-8-BOM | **UTF-16 LE + CRLF + BOM** | `LoadFromText 2` |
| 宏 `.macro` | 无 BOM（纯 ASCII） | 原样 | `LoadFromText 4` |

编码错了**不会报编码错误**，而是报业务错误（中文乱码 → SQL `3075 语法错误(操作符丢失)`）。

## 组件名从哪来（实测）

`VBComponents.Import` 建出的组件名取自**文件头部的 `Attribute VB_Name`**，**不是文件名**：

- `class-module-stub.cls`（头部 `Attribute VB_Name = "clsTemplate"`）导入后 → 组件名 `clsTemplate`；
- `standard-module-stub.bas`（头部 `"basTemplate"`）导入后 → 组件名 `basTemplate`。

推论（重建脚本必须遵守）：
1. **移除旧组件要按 `Attribute VB_Name` 匹配**，按文件名匹配在两者不一致时会漏删 → 留下重名/僵尸模块。
2. 读回比对前必须先剥掉 `VERSION / BEGIN / END / Attribute / MultiUse` 头部行，否则误报"不一致"。

## 重建模块的唯一推荐路径

```bat
python scripts/rebuild_module_from_src.py "<你的库.accdb>" "<要重建的 src 文件>"
```

`rebuild_module_from_src.py` 会自动：读 `.src`（兼容有无 BOM）→ 转 GBK 无 BOM + CRLF
→ `Remove` 同名组件 → `Import` → 读回逐行比对（自动剥 VBE 头部行）→ `DoCmd.Save 5`。

**禁止**再把 `InsertLines` 当成"重建模块"的手段（会部分写入、留残片）；
**禁止**走 `VBComponents.Add(2)` + 改名（改名抛错误 53）。

## 统计真实对象（去噪）

- 真实查询 = `[q for q in QueryDefs if not q.Name.startswith("~")]`（`~sq_` 是 Access 为内嵌 RowSource 自动生成的）。
- 真实表 = `[t for t in TableDefs if not t.Name.startswith("MSys")]`。
- 测"模块存在"必须用 `CurrentProject.AllModules`（`Application.Modules` 只含已打开的模块）。

## VCS 往返验证清单

- [ ] `Application.SaveAsText` 逐类对象导出成功（模块 acModule=5、宏 acMacro=4、窗体 acForm=2）。
- [ ] 导出目录对象数量与实时库一致，无"一边有一边无"。
- [ ] 导入后 `CurrentProject.AllModules` 含期望的类模块且 `c.Type=2`。
- [ ] 类模块可实例化；业务测试 0 失败。
- [ ] 关库重开二次复验：对象仍在、类型不变、可编译。
