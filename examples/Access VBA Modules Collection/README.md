# AccessDevelop

Access VBA 常用模块集合，面向企业级 Access 开发场景，提供可复用的导出、验证、翻译与数据库操作能力。

> 本仓库汇总了在真实项目中反复使用的 VBA 模块与类文件，目标是减少重复造轮子，帮助开发者更快交付稳定的 Access 解决方案。

[中文](#-功能特性) | [English](#-features)

## 徽章

![VBA](https://img.shields.io/badge/VBA-Access-blue)
![Microsoft Access](https://img.shields.io/badge/Microsoft%20Access-2010%2B-red)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ 功能特性

- 提供 Access 常见业务开发所需的模块集合（Excel、PPT、HTML 导出，翻译，验证，ADO 操作）
- 所有模块均可直接导入 `.bas` / `.cls` 到现有 Access 工程使用
- 包含自动编号、字段校验、验证 UI 联动等高频业务能力
- 提供 VBE 工具模块，支持批量为过程注入标准错误处理
- 文档化的模块说明和使用示例，便于二次开发

## 📋 环境要求

| 依赖项 | 版本/要求 | 说明 |
| --- | --- | --- |
| Windows | Windows 10/11（建议） | 运行 Access 与 Office 自动化 |
| Microsoft Access | 2010 或更高版本 | VBA 模块运行环境 |
| Microsoft Excel | 2010 或更高版本 | `modExportToExcel`、`modPasteDataToExcel`、`basExportChart` 相关功能 |
| Microsoft PowerPoint | 2010 或更高版本 | `modExportToPPT` 功能 |
| Azure Translator 服务 | 有效订阅与密钥 | `modAzureTranslator` 翻译能力 |
| VBE Extensibility | 5.3 引用 + 启用信任访问 | `modVBETools` 代码注入能力 |

## 🚀 快速开始

1. 克隆仓库。

```bash
git clone https://github.com/miaowei2/accessdevelop.git
```

2. 将所需模块导入你的 Access 工程（`Alt + F11` -> `文件` -> `导入文件`）。

3. 在 `工具` -> `引用` 中按需勾选：
- `Microsoft PowerPoint XX.0 Object Library`（PPT 导出）
- `Microsoft Visual Basic for Applications Extensibility 5.3`（VBE 工具）

4. 按模块调用入口函数。

```vba
' 智能编号
orderNo = AutoNumStr("订单表", "订单号", 5, "ORD", "yyyymmdd")

' 导出到 Excel
ExportToExcel "我的报表", "销售数据", "A1"

' 数据粘贴并格式化到 Excel
PasteDataToExcel "销售报表", "2024数据", RGB(68, 114, 196), True

' 导出 Access 查询到 PowerPoint
CreateCompleteReport

' Azure 翻译
result = TranslateText("你好，世界", "zh-Hans", "en")
```

## 📁 项目结构

```text
accessdevelop/
|-- ADOExecute.cls
|-- basAutoNumStr.bas
|-- basExportChart.bas
|-- ClsFieldValidator.cls
|-- M_ValidationUI.bas
|-- modAzureTranslator.bas
|-- modExportToExcel.bas
|-- modExportToPPT.bas
|-- modHTMLExport.bas
|-- modPasteDataToExcel.bas
|-- modVBETools.bas
|-- README.md
`-- wiki/
    |-- Changelog.md
    |-- Contributing.md
    |-- FAQ.md
    |-- Home.md
    |-- Installation-Guide.md
    |-- Module-Documentation.md
    `-- Tutorials.md
```

## 🔧 核心模块说明

| 模块/类 | 类型 | 作用 |
| --- | --- | --- |
| `basAutoNumStr.bas` | 标准模块 | 生成可配置前缀与日期的自动编号 |
| `modExportToPPT.bas` | 标准模块 | 将 Access 数据导出为结构化 PowerPoint 报告 |
| `modExportToExcel.bas` | 标准模块 | 导出窗体/数据到 Excel 并进行基础格式处理 |
| `modPasteDataToExcel.bas` | 标准模块 | 粘贴数据到 Excel 并自动美化表头/边框/列宽 |
| `basExportChart.bas` | 标准模块 | 批量导出图表到 Excel 工作簿 |
| `modHTMLExport.bas` | 标准模块 | 将表/查询/窗体/报表导出为 HTML |
| `modAzureTranslator.bas` | 标准模块 | 调用 Azure Translator 实现单条与批量翻译 |
| `ClsFieldValidator.cls` | 类模块 | 提供必填、邮箱、手机号、长度、范围等校验 |
| `M_ValidationUI.bas` | 标准模块 | 提供验证状态 UI 反馈与高亮 |
| `modVBETools.bas` | 标准模块 | 扫描并批量添加错误处理结构 |
| `ADOExecute.cls` | 类模块 | 封装 ADO 连接、执行 SQL、查询与事务 |

常用公开方法（示例）：
- `AutoNumStr(...)`
- `CreateCompleteReport()`
- `ExportToExcel(...)`
- `PasteDataToExcel(...)`
- `ExportChartToExcel(...)`
- `ExportData(...)`
- `TranslateText(...)` / `TranslateMultiple(...)`
- `AddErrorHandlingToAll()`
- `ClsFieldValidator.Validate(...)`
- `UpdateValidationStatus(...)` / `HighlightTextBox(...)`
- `ADOExecute.Connect()` / `ADOExecute.ExecuteSQL(...)` / `ADOExecute.GetValues(...)`

## 🗺️ 路线图

- [x] 提供常用导出模块（Excel/PPT/HTML）
- [x] 提供字段验证与验证 UI 辅助
- [x] 提供 ADO 数据访问封装
- [x] 集成 Azure 翻译调用能力
- [ ] 增加更多标准化错误处理模板
- [ ] 增加更多单元测试/示例数据库
- [ ] 完善每个模块的参数清单与边界条件文档

## 🐛 问题反馈

如果发现 Bug 或有改进建议，请：

- 提交 [Issue](https://github.com/miaowei2/accessdevelop/issues)
- 详细描述问题或建议
- 如可能，提供复现步骤

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 👨‍💻 作者

**缪炜（will miao）**

现任微软最有价值专家（MVP），自媒体博主（公众号Access开发）

微软官方MVP主页地址：[@MVP](https://mvp.microsoft.com/zh-CN/MVP/profile/15c78eb8-1d9d-42de-9c15-afba24ec931d)
拥有丰富的企业级开发与培训经验，曾服务多家外企及知名合资企业，包括：麦格纳电子 (Magna)、飞利浦电子 (Philips)、卡特彼勒 (Caterpillar)、硕腾 (Zoetis)等。

项目经验：深耕企业数字化解决方案，通过 Access 独立架构或者其他语言开发过 ERP（企业资源计划）、WMS（仓储管理）、MES（生产执行）、CRM（客户关系）及 HR 等核心业务系统，具备极强的实战落地能力。熟悉：VBA、C#、JavaScript、SQL等开发语言。

## 📮 联系方式

- GitHub: [@miaowei2](https://github.com/miaowei2)
- email:will.miao@edonsoft.com
- 公众号：Access开发
- B站：[@Access开发易登软件](https://space.bilibili.com/10580232?spm_id_from=333.1007.0.0)
- 公司网站：[www.edonsoft.com](http://www.edonsoft.com)

## 🙏 致谢

感谢所有使用和贡献本项目的开发者！

---

# English

## AccessDevelop

A collection of practical Access VBA modules for enterprise use cases, covering export, validation, translation, and database operations.

> This repository packages reusable VBA modules and class files that have been used repeatedly in real-world projects, helping developers deliver stable Access solutions faster.

![VBA](https://img.shields.io/badge/VBA-Access-blue)
![Microsoft Access](https://img.shields.io/badge/Microsoft%20Access-2010%2B-red)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- A practical toolbox for common Access business development scenarios
- Ready-to-import `.bas` and `.cls` modules for existing Access projects
- Core capabilities including auto numbering, field validation, and validation UI feedback
- VBE helper tooling to batch-insert standardized error handling blocks
- Module-oriented documentation and usage examples for extension

## 📋 Requirements

| Dependency | Version/Requirement | Notes |
| --- | --- | --- |
| Windows | Windows 10/11 (recommended) | Host environment for Office automation |
| Microsoft Access | 2010 or newer | Runtime for VBA modules |
| Microsoft Excel | 2010 or newer | Required by Excel export related modules |
| Microsoft PowerPoint | 2010 or newer | Required by `modExportToPPT` |
| Azure Translator Service | Valid subscription and API key | Required by `modAzureTranslator` |
| VBE Extensibility | 5.3 reference + trusted VBA project access | Required by `modVBETools` |

## 🚀 Quick Start

1. Clone the repository.

```bash
git clone https://github.com/miaowei2/accessdevelop.git
```

2. Import required modules in Access (`Alt + F11` -> `File` -> `Import File`).

3. In `Tools` -> `References`, enable required object libraries as needed.

4. Call entry procedures from your form/module code.

```vba
' Smart numbering
orderNo = AutoNumStr("OrderTable", "OrderNo", 5, "ORD", "yyyymmdd")

' Export to Excel
ExportToExcel "MyReport", "SalesData", "A1"

' Paste and format data in Excel
PasteDataToExcel "Sales Report", "2024 Data", RGB(68, 114, 196), True

' Export Access query data to PowerPoint
CreateCompleteReport

' Azure translation
result = TranslateText("Hello, world", "en", "zh-Hans")
```

## 📁 Project Structure

```text
accessdevelop/
|-- ADOExecute.cls
|-- basAutoNumStr.bas
|-- basExportChart.bas
|-- ClsFieldValidator.cls
|-- M_ValidationUI.bas
|-- modAzureTranslator.bas
|-- modExportToExcel.bas
|-- modExportToPPT.bas
|-- modHTMLExport.bas
|-- modPasteDataToExcel.bas
|-- modVBETools.bas
|-- README.md
`-- wiki/
    |-- Changelog.md
    |-- Contributing.md
    |-- FAQ.md
    |-- Home.md
    |-- Installation-Guide.md
    |-- Module-Documentation.md
    `-- Tutorials.md
```

## 🔧 Core Module Reference

| Module/Class | Type | Purpose |
| --- | --- | --- |
| `basAutoNumStr.bas` | Module | Generate auto numbers with configurable prefix/date patterns |
| `modExportToPPT.bas` | Module | Export Access data to structured PowerPoint reports |
| `modExportToExcel.bas` | Module | Export form/data content to Excel with basic formatting |
| `modPasteDataToExcel.bas` | Module | Paste and beautify data in Excel |
| `basExportChart.bas` | Module | Batch export chart controls to Excel worksheets |
| `modHTMLExport.bas` | Module | Export table/query/form/report to HTML |
| `modAzureTranslator.bas` | Module | Integrate Azure Translator for single/batch translation |
| `ClsFieldValidator.cls` | Class | Validate required/email/mobile/length/range fields |
| `M_ValidationUI.bas` | Module | Provide validation status UI feedback |
| `modVBETools.bas` | Module | Scan and inject standardized error handling blocks |
| `ADOExecute.cls` | Class | Encapsulate ADO connect/execute/query/transaction flow |

Public APIs (examples):
- `AutoNumStr(...)`
- `CreateCompleteReport()`
- `ExportToExcel(...)`
- `PasteDataToExcel(...)`
- `ExportChartToExcel(...)`
- `ExportData(...)`
- `TranslateText(...)` / `TranslateMultiple(...)`
- `AddErrorHandlingToAll()`
- `ClsFieldValidator.Validate(...)`
- `UpdateValidationStatus(...)` / `HighlightTextBox(...)`
- `ADOExecute.Connect()` / `ADOExecute.ExecuteSQL(...)` / `ADOExecute.GetValues(...)`

## 🗺️ Roadmap

- [x] Core export modules (Excel/PPT/HTML)
- [x] Field validation and validation UI helpers
- [x] ADO data access encapsulation
- [x] Azure Translator integration
- [ ] More standardized error-handling templates
- [ ] More test cases and sample databases
- [ ] Complete parameter and edge-case documentation per module

## 🐛 Bug Reports

If you find a bug or have a suggestion:

- Submit an [Issue](https://github.com/miaowei2/accessdevelop/issues)
- Describe the problem or suggestion in detail
- Include reproduction steps if possible

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## 👨‍💻 Author

**Will Miao (缪炜)**

Microsoft Most Valuable Professional (MVP) and content creator (WeChat Official Account: Access开发).

Microsoft MVP profile: [@MVP](https://mvp.microsoft.com/zh-CN/MVP/profile/15c78eb8-1d9d-42de-9c15-afba24ec931d)

Extensive experience in enterprise-level development and training, having served multinational and joint-venture companies including Magna Electronics, Philips, Caterpillar, Zoetis, among others.

Specializes in enterprise digital solutions, having independently architected or co-developed ERP, WMS, MES, CRM, and HR systems using Access and other technologies. Proficient in VBA, C#, JavaScript, SQL, and more.

## 📮 Contact

- GitHub: [@miaowei2](https://github.com/miaowei2)
- Email: will.miao@edonsoft.com
- WeChat Official Account: Access开发
- Bilibili: [@Access开发易登软件](https://space.bilibili.com/10580232?spm_id_from=333.1007.0.0)
- Website: [www.edonsoft.com](http://www.edonsoft.com)

## 🙏 Acknowledgments

Thanks to all developers who use and contribute to this project!
