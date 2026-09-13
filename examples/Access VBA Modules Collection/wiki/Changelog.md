# 更新日志 | Changelog

所有重要更改都将记录在此文件中。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [未发布]

### 计划中
- modExportToWord - Word 文档导出模块
- basDataValidator - 数据完整性验证
- modEmailSender - 邮件发送模块
- clsLogger - 日志记录类

---

## [1.0.0] - 2026-01-23

### 新增
- ✨ 初始发布
- 📦 包含 11 个实用 VBA 模块
- 📝 完整的中英文文档
- 🌐 GitHub Wiki 文档站点

### 模块列表

#### 数据导出模块
- `modExportToPPT` - PowerPoint 报告导出
- `modExportToExcel` - Excel 数据导出
- `modPasteDataToExcel` - 智能 Excel 粘贴
- `basExportChart` - 图表批量导出
- `modHTMLExport` - HTML 格式导出

#### 数据处理模块
- `basAutoNumStr` - 智能自动编号生成器
- `ADOExecute` - ADO 数据库操作封装

#### 开发工具模块
- `modVBETools` - VBA 代码工具（批量错误处理）
- `modAzureTranslator` - Azure 翻译服务集成

#### 验证模块
- `ClsFieldValidator` - 字段验证类
- `M_ValidationUI` - 验证界面辅助模块

### 文档
- 📖 [安装指南](Installation-Guide)
- 📚 [模块文档](Module-Documentation)
- 🎓 [使用教程](Tutorials)
- ❓ [常见问题](FAQ)
- 🤝 [贡献指南](Contributing)

---

## 版本说明

### 版本号规则

使用语义化版本号：`主版本号.次版本号.修订号`

- **主版本号**：不兼容的 API 变更
- **次版本号**：向后兼容的功能新增
- **修订号**：向后兼容的问题修复

### 变更类型

- `新增` - 新功能或模块
- `变更` - 现有功能的变更
- `废弃` - 即将移除的功能
- `移除` - 已移除的功能
- `修复` - Bug 修复
- `安全` - 安全问题修复

---

## 路线图

### v1.1.0（计划中）
- [ ] modExportToWord 模块
- [ ] 增强 modExportToPPT 的图表支持
- [ ] 优化 basAutoNumStr 并发性能
- [ ] 添加更多验证规则到 ClsFieldValidator

### v1.2.0（计划中）
- [ ] modEmailSender 邮件发送模块
- [ ] clsLogger 日志记录类
- [ ] modBackup 数据库备份工具
- [ ] 性能优化

### v2.0.0（远期规划）
- [ ] 重构核心架构
- [ ] 插件系统
- [ ] 图形化配置工具
- [ ] 更多集成选项

---

## 贡献者

感谢以下贡献者（按时间顺序）：

- [@miaowei2](https://github.com/miaowei2) - 项目创建者和主要维护者

*想出现在这里？查看 [贡献指南](Contributing)*

---

## 支持

- 🐛 [报告 Bug](https://github.com/miaowei2/accessdevelop/issues)
- 💡 [功能建议](https://github.com/miaowei2/accessdevelop/issues)
- 📧 Email: will.miao@edonsoft.com

---

**注意**：此更新日志从 v1.0.0 开始维护。

[未发布]: https://github.com/miaowei2/accessdevelop/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/miaowei2/accessdevelop/releases/tag/v1.0.0
