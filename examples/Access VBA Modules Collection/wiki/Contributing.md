# 贡献指南 | Contributing Guide

感谢您考虑为 Access VBA 常用模块集合做贡献！

## 🎯 贡献方式

我们欢迎各种形式的贡献：

- 🐛 **报告 Bug** - 发现问题请及时反馈
- 💡 **提出建议** - 新功能或改进想法
- 📝 **完善文档** - 修正错误、补充说明
- 💻 **贡献代码** - 提交新模块或改进现有代码
- 🌍 **翻译文档** - 帮助翻译成其他语言
- ⭐ **分享项目** - 让更多人了解这个项目

---

## 📋 行为准则

### 我们的承诺

为了营造开放和友好的环境，我们承诺：

- ✅ 使用友好和包容的语言
- ✅ 尊重不同的观点和经验
- ✅ 优雅地接受建设性批评
- ✅ 关注对社区最有利的事情
- ✅ 对其他社区成员表示同理心

### 不可接受的行为

- ❌ 使用性化的语言或图像
- ❌ 人身攻击或侮辱性评论
- ❌ 骚扰行为（公开或私下）
- ❌ 未经许可发布他人的私人信息
- ❌ 其他在专业环境中不当的行为

---

## 🐛 报告 Bug

### 提交 Bug 前

1. **检查是否已存在**
   - 搜索 [Issues](https://github.com/miaowei2/accessdevelop/issues)
   - 查看 [FAQ](FAQ)

2. **确认是 Bug**
   - 在干净的环境中重现问题
   - 排除本地配置问题

### Bug 报告模板

提交 Issue 时请包含：

```markdown
## Bug 描述
简要描述问题

## 复现步骤
1. 打开...
2. 点击...
3. 看到错误...

## 期望行为
描述您期望发生什么

## 实际行为
描述实际发生了什么

## 截图
如果可能，添加截图

## 环境信息
- OS: [例如 Windows 10]
- Access 版本: [例如 Access 2019]
- Office 位数: [32位/64位]
- 模块版本: [例如 v1.0.0]

## 错误信息
```vba
' 粘贴完整的错误消息
```

## 附加信息
其他可能有帮助的信息
```

---

## 💡 功能建议

### 提交建议前

1. **确认需求**
   - 这个功能对多数用户有用吗？
   - 是否可以通过现有模块组合实现？

2. **检查是否已存在**
   - 搜索现有 Issues
   - 查看项目路线图

### 建议模板

```markdown
## 功能描述
清晰简洁地描述您想要的功能

## 使用场景
为什么需要这个功能？它解决什么问题？

## 建议的解决方案
您认为应该如何实现？

## 替代方案
您考虑过的其他方案

## 附加信息
其他相关信息、示例或参考
```

---

## 💻 贡献代码

### 开发环境设置

1. **Fork 仓库**
   - 访问 [GitHub 仓库](https://github.com/miaowei2/accessdevelop)
   - 点击右上角 "Fork" 按钮

2. **克隆仓库**
   ```bash
   git clone https://github.com/你的用户名/accessdevelop.git
   cd accessdevelop
   ```

3. **创建分支**
   ```bash
   git checkout -b feature/你的功能名称
   ```

### 代码规范

#### 命名约定

```vba
' 模块名称
' - 标准模块：mod前缀，例如 modExportToExcel
' - 类模块：Cls前缀，例如 ClsFieldValidator
' - 基础模块：bas前缀，例如 basAutoNumStr

' 函数和过程
' - 使用 PascalCase（每个单词首字母大写）
Public Function ExportToExcel()
Public Sub CreateReport()

' 变量
' - 局部变量：camelCase（首字母小写）
Dim fileName As String
Dim recordCount As Long

' - 私有成员变量：m_前缀 + camelCase
Private m_connectionString As String
Private m_errorMessage As String

' 常量
' - 全大写，下划线分隔
Public Const MAX_RECORDS = 1000
Private Const ERROR_MESSAGE = "发生错误"
```

#### 代码格式

```vba
' 1. 缩进：使用 Tab（4 个空格）
Public Sub Example()
    If condition Then
        ' 代码
    End If
End Sub

' 2. 每行不超过 100 字符
' 3. 使用空行分隔逻辑块
' 4. 运算符两侧加空格
result = value1 + value2

' 5. 逗号后加空格
Function Calculate(param1 As String, param2 As Long)
```

#### 注释规范

```vba
' ==================== 模块说明 ====================
' 模块名称：modExportToExcel
' 功能：将 Access 数据导出到 Excel
' 作者：你的名字
' 日期：2026-01-23
' ====================================================

' 函数说明
' 参数：
'   - tableName: 表名称
'   - outputPath: 输出路径
' 返回：
'   - Boolean: 成功返回 True
Public Function ExportTable(tableName As String, _
                           outputPath As String) As Boolean
    ' 实现代码
End Function
```

#### 错误处理

```vba
' 所有公共函数必须包含错误处理
Public Function YourFunction() As Boolean
    On Error GoTo ErrorHandler
    
    ' 函数实现
    
    YourFunction = True
    Exit Function
    
ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    YourFunction = False
End Function
```

### 提交代码

#### Commit 信息规范

使用清晰的 commit 信息：

```bash
# 格式：<类型>: <简短描述>

# 类型：
# - feat: 新功能
# - fix: 修复 Bug
# - docs: 文档更新
# - style: 代码格式（不影响功能）
# - refactor: 重构
# - test: 测试相关
# - chore: 构建或辅助工具

# 示例
feat: 添加 Excel 导出模块
fix: 修复自动编号重复问题
docs: 更新安装指南
```

#### 提交 Pull Request

1. **推送分支**
   ```bash
   git push origin feature/你的功能名称
   ```

2. **创建 Pull Request**
   - 访问您的 Fork 页面
   - 点击 "New Pull Request"
   - 填写 PR 描述

3. **PR 描述模板**
   ```markdown
   ## 变更说明
   简要描述这个 PR 的内容
   
   ## 变更类型
   - [ ] Bug 修复
   - [ ] 新功能
   - [ ] 文档更新
   - [ ] 代码重构
   
   ## 测试
   描述如何测试这些变更
   
   ## 相关 Issue
   Fixes #issue编号
   
   ## 检查清单
   - [ ] 代码遵循项目规范
   - [ ] 添加了必要的注释
   - [ ] 更新了相关文档
   - [ ] 通过了测试
   - [ ] 没有引入新的警告
   ```

---

## 📝 贡献文档

### 文档结构

```
wiki/
├── Home.md                    # Wiki 首页
├── Installation-Guide.md      # 安装指南
├── Module-Documentation.md    # 模块总览
├── Tutorials.md              # 使用教程
├── FAQ.md                    # 常见问题
├── Contributing.md           # 本文档
├── Changelog.md              # 更新日志
└── [模块名].md               # 各模块详细文档
```

### 文档规范

1. **Markdown 格式**
   - 使用标准 Markdown 语法
   - 代码块指定语言（vba, bash 等）

2. **标题层级**
   ```markdown
   # 一级标题（页面标题）
   ## 二级标题（主要章节）
   ### 三级标题（小节）
   #### 四级标题（细节）
   ```

3. **代码示例**
   ```markdown
   ```vba
   ' VBA 代码示例
   Public Sub Example()
       Debug.Print "Hello"
   End Sub
   ```​
   ```

4. **链接**
   ```markdown
   [链接文字](URL)
   [模块文档](Module-Documentation)
   ```

---

## 🧪 测试要求

### 新模块提交

提交新模块时必须包含：

1. **测试用例**
   ```vba
   ' 在模块注释中说明测试方法
   ' 测试用例：
   ' 1. 正常情况：...
   ' 2. 边界情况：...
   ' 3. 异常情况：...
   ```

2. **使用示例**
   - 提供至少 3 个使用示例
   - 覆盖常见使用场景

3. **依赖说明**
   - 列出所需的对象库引用
   - 说明系统要求

### Bug 修复

修复 Bug 时：

1. 添加防止回归的测试
2. 说明如何复现原 Bug
3. 说明修复方案

---

## 🎁 新模块提交清单

提交新模块前请确认：

- [ ] **代码质量**
  - [ ] 遵循命名规范
  - [ ] 包含完整的错误处理
  - [ ] 添加了详细注释
  - [ ] 代码格式规范

- [ ] **文档**
  - [ ] 模块顶部有完整说明
  - [ ] 每个公共函数有注释
  - [ ] 添加使用示例
  - [ ] 更新 README.md

- [ ] **测试**
  - [ ] 在 Access 2010+ 中测试
  - [ ] 测试正常和异常情况
  - [ ] 验证错误处理
  - [ ] 检查内存泄漏（对象释放）

- [ ] **兼容性**
  - [ ] 32 位和 64 位 Office
  - [ ] Access 2010 及以上版本
  - [ ] 不依赖外部 DLL（除 Office 自带）

---

## 🏆 贡献者

感谢所有为项目做出贡献的开发者！

您的贡献将被记录在：
- 项目 README
- 发布说明
- 贡献者列表

---

## 📞 联系方式

有任何问题或建议？

- **GitHub Issues**: [提交问题](https://github.com/miaowei2/accessdevelop/issues)
- **Email**: will.miao@edonsoft.com
- **公众号**: Access开发
- **B站**: [@Access开发易登软件](https://space.bilibili.com/10580232)

---

## 📄 许可证

贡献代码即表示您同意：

1. 您的贡献将使用项目的 MIT 许可证
2. 您拥有贡献代码的权利
3. 您的贡献不侵犯任何第三方权利

---

**再次感谢您的贡献！** 🎉

每一个贡献，无论大小，都让这个项目变得更好！
