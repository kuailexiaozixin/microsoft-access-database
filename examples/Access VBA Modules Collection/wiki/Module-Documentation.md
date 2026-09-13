# 模块文档 | Module Documentation

本页面提供所有模块的快速参考索引。点击模块名称查看详细文档。

## 📚 模块分类

### 🎨 数据导出模块

| 模块名称 | 功能简介 | 复杂度 |
|---------|---------|--------|
| [modExportToPPT](modExportToPPT) | 将查询数据导出为专业的 PowerPoint 报告 | ⭐⭐⭐ |
| [modExportToExcel](modExportToExcel) | 导出窗体数据到 Excel 并自动格式化 | ⭐⭐ |
| [modPasteDataToExcel](modPasteDataToExcel) | 智能粘贴数据到 Excel（带美化） | ⭐⭐ |
| [basExportChart](basExportChart) | 批量导出图表控件到 Excel | ⭐⭐ |
| [modHTMLExport](modHTMLExport) | 导出 Access 对象为 HTML 格式 | ⭐⭐ |

### 🔢 数据处理模块

| 模块名称 | 功能简介 | 复杂度 |
|---------|---------|--------|
| [basAutoNumStr](basAutoNumStr) | 生成智能自动编号（订单号、单据号等） | ⭐ |
| [ADOExecute](ADOExecute) | ADO 数据库操作封装（事务支持） | ⭐⭐⭐ |

### 🛠️ 开发工具模块

| 模块名称 | 功能简介 | 复杂度 |
|---------|---------|--------|
| [modVBETools](modVBETools) | 批量为代码添加错误处理 | ⭐⭐⭐ |
| [modAzureTranslator](modAzureTranslator) | Azure 翻译服务集成 | ⭐⭐⭐ |

### ✅ 验证模块

| 模块名称 | 功能简介 | 复杂度 |
|---------|---------|--------|
| [ClsFieldValidator](ClsFieldValidator) | 强大的字段验证类（邮箱、手机等） | ⭐⭐ |
| [M_ValidationUI](M_ValidationUI) | 验证结果的可视化反馈 | ⭐ |

---

## 🔍 快速查找

### 按使用场景查找

**我想要...**

- **生成订单编号** → [basAutoNumStr](basAutoNumStr)
- **导出数据到 PowerPoint** → [modExportToPPT](modExportToPPT)
- **导出数据到 Excel** → [modExportToExcel](modExportToExcel) 或 [modPasteDataToExcel](modPasteDataToExcel)
- **导出图表** → [basExportChart](basExportChart)
- **验证邮箱格式** → [ClsFieldValidator](ClsFieldValidator)
- **翻译文本** → [modAzureTranslator](modAzureTranslator)
- **批量添加错误处理** → [modVBETools](modVBETools)
- **执行带事务的 SQL** → [ADOExecute](ADOExecute)

---

## 📖 模块详细列表

### 1. modExportToPPT

**导出查询数据到 PowerPoint**

- 📄 [完整文档](modExportToPPT)
- 🎯 用途：生成专业的数据分析报告
- ⚙️ 需要引用：Microsoft PowerPoint Object Library
- 💡 示例：`CreateCompleteReport`

---

### 2. modExportToExcel

**快速导出窗体数据到 Excel**

- 📄 [完整文档](modExportToExcel)
- 🎯 用途：导出当前窗体显示的数据
- ⚙️ 需要引用：Microsoft Excel Object Library
- 💡 示例：`ExportToExcel "销售报表"`

---

### 3. modPasteDataToExcel

**智能 Excel 粘贴（带格式化）**

- 📄 [完整文档](modPasteDataToExcel)
- 🎯 用途：粘贴数据并自动美化
- ⚙️ 需要引用：Microsoft Excel Object Library
- 💡 示例：`PasteDataToExcel "报表", "数据", RGB(68, 114, 196), True`

---

### 4. basExportChart

**批量导出图表到 Excel**

- 📄 [完整文档](basExportChart)
- 🎯 用途：导出窗体中的多个图表
- ⚙️ 需要引用：Microsoft Excel Object Library
- 💡 示例：`ExportChartToExcel Array("图表1", "图表2"), "C:\Charts.xlsx"`

---

### 5. modHTMLExport

**导出为 HTML 格式**

- 📄 [完整文档](modHTMLExport)
- 🎯 用途：将表、查询等导出为网页
- ⚙️ 需要引用：无
- 💡 示例：`ExportData "Table", "客户表", "C:\output.html"`

---

### 6. basAutoNumStr

**智能自动编号生成器**

- 📄 [完整文档](basAutoNumStr)
- 🎯 用途：生成订单号、单据号等
- ⚙️ 需要引用：无
- 💡 示例：`AutoNumStr("订单表", "订单号", 5, "ORD", "yyyymmdd")`

---

### 7. ADOExecute

**ADO 数据库操作类**

- 📄 [完整文档](ADOExecute)
- 🎯 用途：执行带事务的 SQL 操作
- ⚙️ 需要引用：Microsoft ActiveX Data Objects Library
- 💡 示例：`db.ExecuteSQL "INSERT INTO ..."`

---

### 8. modVBETools

**VBA 开发工具**

- 📄 [完整文档](modVBETools)
- 🎯 用途：批量添加错误处理代码
- ⚙️ 需要引用：VBA Extensibility 5.3
- 💡 示例：`AddErrorHandlingToAll`
- ⚠️ 注意：需要启用"信任 VBA 工程对象模型"

---

### 9. modAzureTranslator

**Azure 翻译服务**

- 📄 [完整文档](modAzureTranslator)
- 🎯 用途：多语言文本翻译
- ⚙️ 需要引用：Microsoft XML v6.0
- 💡 示例：`TranslateText("Hello", "en", "zh-Hans")`
- ⚠️ 注意：需要配置 Azure 密钥

---

### 10. ClsFieldValidator

**字段验证类**

- 📄 [完整文档](ClsFieldValidator)
- 🎯 用途：验证邮箱、手机号、身份证等
- ⚙️ 需要引用：无
- 💡 示例：`validator.Validate(Me.txtEmail, vtEmail)`

---

### 11. M_ValidationUI

**验证界面辅助**

- 📄 [完整文档](M_ValidationUI)
- 🎯 用途：可视化显示验证结果
- ⚙️ 需要引用：无
- 💡 示例：`UpdateValidationStatus Me.lblStatus, vrValid`
- 💡 配合 ClsFieldValidator 使用效果更佳

---

## 📊 模块依赖关系

```
ClsFieldValidator ←→ M_ValidationUI (推荐配合使用)

modExportToPPT → PowerPoint Object Library
modExportToExcel → Excel Object Library
modPasteDataToExcel → Excel Object Library
basExportChart → Excel Object Library
modVBETools → VBA Extensibility
modAzureTranslator → XML Library + Azure Account
ADOExecute → ADO Library
```

---

## 🆕 版本说明

所有模块都在持续更新中。查看 [Changelog](Changelog) 了解最新变化。

---

## 📝 贡献新模块

如果您有好的模块想要分享，请查看 [贡献指南](Contributing)。

---

**下一步**：
- 选择一个模块查看详细文档
- 查看 [使用教程](Tutorials) 学习实用技巧
- 遇到问题？访问 [FAQ](FAQ)
