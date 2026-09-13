# VCS 文本源编辑规则（改 .src 文本的硬规则）

> 本文件是主技能自有文件，沉淀「改 `.src` 文本」的通用硬规则——AI 文本化开发（导出 → 改文本 → 构建）是本技能工作流的默认改库方式，这些规则是**每个阶段改文本前必读**的操作口径。按对象类型查阅更详细手册时，见文末映射表（手册位于本技能 `references/`，与本文平级，内容不重复：本文讲通用规则，手册讲对象级格式细节）。

违反本文件规则会静默损坏库，或在回写时报 `Error 2128`（LoadComponentFromText 导入失败）。

## 一、VCS 工程结构

一个 VCS 导出的 Access 工程位于 `<数据库名>.accdb.src/` 目录，目录名**必须**与 `.accdb` 文件名一致并追加 `.src`（如 `MyApp.accdb` ↔ `MyApp.accdb.src/`）。工程内：
- 顶层配置文件：`vcs-index.json`（对象清单，删除对象必须同步删条目）、`vcs-options.json`（加载项设置）、`dbs-properties.json`（库属性与启动配置）、`vbe-project.json`（VBA 工程与主文件）、`vbe-references.json`（引用）、`db-connection.json`（链接表连接串）。
- 对象目录：`forms/`、`reports/`、`queries/`、`modules/`、`tbldefs/`、`relations/`、`macros/`（按需）。

## 二、对象文件配对规则

配对规则以 `SKILL.md`「二、对象全景与文本形态」为**单点定义**（11 类对象各自的文件形态：窗体/报表 `.bas`+可选 `.cls`+可选 `.json`、查询 `.bas`+`.sql`、模块 `.bas`、宏 `.bas`、链接表 `.json`、本地表 `.xml`、关系 `.json`、共享图像 `.json`+图像、导入导出规格 `.json`、主题 `.thmx` 等），此处不重复列表。本文只强调两条操作要点：

- **查询的 `.bas`+`.sql` 恒成对**（`.bas`=元数据/设计器状态，`.sql`=SQL 文本）。
- 删除对象时，必须**连它的全部配对文件一起删**，并删除 `vcs-index.json` 中对应条目。

## 三、编码与换行（最易踩坑）

- 所有 JSON 配置文件与 `.sql` 文件用 **UTF-8 带 BOM**（字节 `EF BB BF`）；`.bas`（窗体/报表/宏）、`.cls`、`tbldefs`/`relations` 下文件额外用 **CRLF** 换行（用 `.gitattributes` 强制）。
- **新建文件**时，常规写入工具（Write/echo/cat）默认产出 UTF-8 无 BOM + LF——会损坏导入。必须显式产出字节：

```bash
printf '\xEF\xBB\xBF{\r\n  "Info": { ... }\r\n}\r\n' > path/to/new-file.json
```

- 写后验证：`od -c -N 5 path/to/new-file.json` 前三个字节应为 `357 273 277`（BOM），且 `\r` 在每个 `\n` 前。
- **编辑已有文件**（Edit/Read 读写）保留原编码，不受上述新建问题影响。
- 主技能的编码铁律（导入 Access 前的 GBK 无 BOM + CRLF / 窗体 UTF-16 LE + BOM）见 `references/vba-encoding-rules.md`——两者是同一件事的两个方向：**.src 内是 UTF-8-BOM；写回 Access 前按对象类型转对应编码**。

## 四、安全删除清单（删任何对象前）

1. **搜引用**：在全部 `.cls`/`.bas`/`.sql` 中搜——VBA 代码（`DoCmd.OpenForm`/`OpenReport`/`OutputTo`）、查询 SQL（`FROM 表名`、子报表 `SourceObject`）、窗体 `RecordSource` 属性、`DLookup` 表达式。
2. **查命名冲突**：相似名称可能服务于完全不同的目的（如一个在用报表和一个废弃集成表共享前缀）。
3. **删全部配对文件**（`.bas`+`.sql`、`.bas`+`.cls`+`.json` 等）。
4. **清配置文件**：`vcs-index.json`、`db-connection.json`、`hidden-attributes.json` 中的对应条目。
5. **重排 TabIndex**：窗体/报表删控件后重排所在节的 TabIndex。
6. **测试导入**：改动后在 Access 里实际导入验证。

## 五、常见坑

- **注释里的引用**：注释掉的 VBA 行若引用你要保留的对象，别动它。
- **db-connection.json**：只能有一个连接条目，且须与所有链接表/传递查询使用的规范 ODBC 连接串一致；绝不写明文口令或用户特定凭据。
- **vcs-index.json 是真相**：导入/导出以它为索引，手工改了对象文件却不同步索引，构建会漏对象。

## 六、按对象类型查手册（导航见 references/README.md）

需要对象级的细节（控件属性、图片、条件格式、工程配置等）时，按类型读对应手册——**9 份手册的导航入口以 `references/README.md`『按问题读』为单点**（窗体/报表/图片/条件格式/查询/表关系/模块/工程配置/宏与图像等一对象一份，独立成文），此处不重复列表。

完整可运行示例：`examples/CustomerOrders/CustomerOrders.accdb.src/`（3 表 4 查询 6 窗体 1 报表的订单系统，一个 prompt 生成）。**阅读法：与 9 份手册逐条对照**——每种对象先在对应手册读格式规则（手册清单见 `references/README.md`『按问题读』），再在示例里找到该对象的实际文件核对落地写法：窗体→`forms/fCustomerList.bas`+`.cls`、报表→`reports/rOrderReport.bas`、查询→`queries/qOrders.bas`+`.sql`、VBA→`modules/modNavigation.bas`、表→`tbldefs/tOrders.xml`、关系→`relations/tCustomerstOrders.json`、工程配置→顶层 6 个 `.json`。规则与活样例并读，AI 生成/修改任何对象前先照此对一遍。
