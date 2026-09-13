# 安装指南 | Installation Guide

本指南将帮助您在 Microsoft Access 项目中安装和配置 VBA 模块。

## 📋 目录

- [系统要求](#系统要求)
- [基础安装步骤](#基础安装步骤)
- [配置对象库引用](#配置对象库引用)
- [验证安装](#验证安装)
- [常见问题](#常见问题)

---

## 系统要求

### 必需软件

- **Microsoft Access** 2010 或更高版本
- **Windows** 7 或更高版本

### 可选软件（根据使用的模块）

- **Microsoft PowerPoint** 2010+ (用于 modExportToPPT)
- **Microsoft Excel** 2010+ (用于导出模块)
- **Internet 连接** (用于 modAzureTranslator)

---

## 基础安装步骤

### 方法一：下载整个仓库

1. **克隆或下载仓库**
   ```bash
   git clone https://github.com/miaowei2/accessdevelop.git
   ```
   
   或直接下载 ZIP 文件：
   - 访问 [GitHub 仓库](https://github.com/miaowei2/accessdevelop)
   - 点击绿色的 "Code" 按钮
   - 选择 "Download ZIP"

2. **解压文件**（如果下载了 ZIP）
   - 将文件解压到本地目录

### 方法二：下载单个模块文件

1. 在 GitHub 上找到您需要的 `.bas` 或 `.cls` 文件
2. 点击文件名打开文件
3. 点击 "Raw" 按钮
4. 右键 → 另存为，保存到本地

---

## 导入模块到 Access

### 导入标准模块 (.bas)

1. **打开您的 Access 数据库**
   
2. **打开 VBA 编辑器**
   - 按快捷键 `Alt + F11`
   - 或通过菜单：数据库工具 → Visual Basic

3. **导入模块文件**
   - 在 VBA 编辑器中，点击菜单：文件 → 导入文件
   - 找到下载的 `.bas` 文件
   - 点击"打开"

4. **验证导入**
   - 在左侧项目资源管理器中，展开"模块"节点
   - 确认模块已成功导入

### 导入类模块 (.cls)

步骤与标准模块相同，导入后模块会出现在"类模块"节点下。

---

## 配置对象库引用

某些模块需要添加外部对象库引用才能正常工作。

### 添加引用的步骤

1. **打开 VBA 编辑器** (`Alt + F11`)

2. **打开引用对话框**
   - 点击菜单：工具 → 引用

3. **添加所需引用**
   - 在弹出的对话框中，向下滚动找到需要的库
   - 勾选复选框
   - 点击"确定"

### 各模块所需引用

#### modExportToPPT
```
☑ Microsoft PowerPoint XX.0 Object Library
```

#### modExportToExcel / modPasteDataToExcel
```
☑ Microsoft Excel XX.0 Object Library
```

#### modVBETools
```
☑ Microsoft Visual Basic for Applications Extensibility 5.3
```

**重要提示**：使用 modVBETools 时，还需要启用以下设置：
- 文件 → 选项 → 信任中心 → 信任中心设置
- 宏设置 → 勾选"信任对 VBA 工程对象模型的访问"

#### ADOExecute
```
☑ Microsoft ActiveX Data Objects 6.x Library
```

#### modAzureTranslator
```
☑ Microsoft XML, v6.0
```

---

## 配置模块参数

某些模块需要配置参数才能使用。

### modAzureTranslator 配置

打开 `modAzureTranslator.bas`，修改以下常量：

```vba
' 在模块顶部找到这些行并修改
Private Const AZURE_KEY As String = "您的Azure订阅密钥"
Private Const AZURE_REGION As String = "您的区域"  ' 例如: "eastus"
Private Const AZURE_ENDPOINT As String = "您的端点地址"
```

### 获取 Azure 翻译服务密钥

1. 访问 [Azure Portal](https://portal.azure.com)
2. 创建"翻译器"资源
3. 获取密钥和端点
4. 将密钥填入模块配置

---

## 验证安装

### 测试模块是否正常工作

1. **打开立即窗口**
   - 在 VBA 编辑器中按 `Ctrl + G`

2. **运行简单测试**

   ```vba
   ' 测试 basAutoNumStr
   ? AutoNumStr("任意表名", "任意字段", 5, "TEST")
   ' 应该返回类似: TEST00001
   ```

3. **如果出现错误**
   - 检查是否添加了所需的对象库引用
   - 检查模块是否完整导入
   - 查看 [常见问题](FAQ) 页面

---

## 最佳实践

### 💡 建议

1. **备份数据库**
   - 在导入新模块前，始终备份您的数据库文件

2. **测试环境**
   - 首先在测试数据库中导入和测试模块
   - 确认无误后再应用到生产环境

3. **版本控制**
   - 记录您使用的模块版本
   - 定期检查更新

4. **代码审查**
   - 导入前查看模块代码
   - 理解模块功能后再使用

---

## 卸载模块

如果需要移除某个模块：

1. 打开 VBA 编辑器 (`Alt + F11`)
2. 在项目资源管理器中找到模块
3. 右键点击模块名称
4. 选择"移除 [模块名]"
5. 选择是否导出备份（建议选"是"）

---

## 下一步

- 查看 [模块文档](Module-Documentation) 了解各模块详细用法
- 阅读 [使用教程](Tutorials) 学习实用技巧
- 遇到问题？查看 [常见问题](FAQ)

---

**有问题？** 请访问 [GitHub Issues](https://github.com/miaowei2/accessdevelop/issues) 提问。
