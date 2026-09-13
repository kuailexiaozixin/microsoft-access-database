# 使用教程 | Tutorials

本页面提供各种常见场景的详细教程和最佳实践。

## 📚 教程目录

- [入门教程](#入门教程)
- [数据导出教程](#数据导出教程)
- [数据验证教程](#数据验证教程)
- [高级技巧](#高级技巧)

---

## 入门教程

### 🎯 教程 1：生成订单编号

**场景**：为新订单自动生成唯一的订单编号

**步骤**：

1. **导入模块**
   ```
   导入 basAutoNumStr.bas
   ```

2. **在窗体的"新记录前"事件中使用**
   ```vba
   Private Sub Form_BeforeInsert(Cancel As Integer)
       ' 生成格式如：ORD20260123001
       Me.订单号 = AutoNumStr("订单表", "订单号", 3, "ORD", "yyyymmdd")
   End Sub
   ```

3. **测试**
   - 打开窗体
   - 点击"新记录"
   - 订单号自动填充

**结果**：
```
第一条记录：ORD20260123001
第二条记录：ORD20260123002
第三条记录：ORD20260123003
```

---

### 🎯 教程 2：快速导出数据到 Excel

**场景**：将当前查看的数据快速导出到 Excel

**步骤**：

1. **导入模块**
   ```
   导入 modExportToExcel.bas
   ```

2. **添加对象库引用**
   - 工具 → 引用
   - 勾选 "Microsoft Excel XX.0 Object Library"

3. **添加导出按钮**
   - 在窗体上添加一个命令按钮
   - 命名为 `btnExport`

4. **编写按钮点击事件**
   ```vba
   Private Sub btnExport_Click()
       ' 基础用法
       ExportToExcel "销售数据报表"
       
       ' 或指定工作表名称
       ' ExportToExcel "销售数据报表", "2024年度"
   End Sub
   ```

5. **测试**
   - 在窗体中查看数据
   - 点击导出按钮
   - 选择保存位置
   - Excel 自动打开并显示数据

---

## 数据导出教程

### 🎯 教程 3：创建专业的 PowerPoint 报告

**场景**：每周自动生成销售分析 PPT

**步骤**：

1. **准备查询**
   - 创建查询：`qry销售统计`
   - 创建查询：`qry客户分析`
   - 创建查询：`qry产品排名`

2. **导入模块**
   ```
   导入 modExportToPPT.bas
   ```

3. **添加引用**
   ```
   Microsoft PowerPoint XX.0 Object Library
   ```

4. **修改模块代码**
   ```vba
   ' 在 CreateCompleteReport 函数中修改
   Public Sub CreateCompleteReport()
       ' ... 前面的代码 ...
       
       ' 添加数据幻灯片
       Call AddQuerySlide(pptPres, "qry销售统计", "销售数据分析", 3)
       Call AddQuerySlide(pptPres, "qry客户分析", "客户分布情况", 4)
       Call AddQuerySlide(pptPres, "qry产品排名", "产品销售排名", 5)
       
       ' ... 后面的代码 ...
   End Sub
   ```

5. **创建按钮执行**
   ```vba
   Private Sub btn生成报告_Click()
       CreateCompleteReport
       MsgBox "报告已生成！", vbInformation
   End Sub
   ```

**高级技巧**：
- 在模块中修改配色方案
- 自定义封面页标题
- 调整表格样式

---

### 🎯 教程 4：批量导出图表

**场景**：将多个数据分析图表导出到一个 Excel 文件

**步骤**：

1. **准备图表控件**
   - 窗体上有多个图表：`chart销售趋势`、`chart地区分布`、`chart产品对比`

2. **导入模块**
   ```
   导入 basExportChart.bas
   ```

3. **编写导出代码**
   ```vba
   Private Sub btn导出图表_Click()
       Dim chartNames As Variant
       Dim savePath As String
       
       ' 定义要导出的图表
       chartNames = Array("chart销售趋势", "chart地区分布", "chart产品对比")
       
       ' 设置保存路径
       savePath = CurrentProject.Path & "\图表分析_" & Format(Date, "yyyymmdd") & ".xlsx"
       
       ' 执行导出
       If ExportChartToExcel(chartNames, savePath) Then
           MsgBox "图表导出成功！" & vbCrLf & savePath, vbInformation
       Else
           MsgBox "导出失败！", vbCritical
       End If
   End Sub
   ```

---

## 数据验证教程

### 🎯 教程 5：实现表单验证

**场景**：用户注册表单，验证邮箱和手机号

**步骤**：

1. **导入模块**
   ```
   导入 ClsFieldValidator.cls
   导入 M_ValidationUI.bas
   ```

2. **设计窗体**
   ```
   文本框：txtEmail（邮箱）
   标签：lblEmailStatus（验证状态）
   文本框：txtPhone（手机号）
   标签：lblPhoneStatus（验证状态）
   按钮：btn提交
   ```

3. **编写验证代码**
   ```vba
   ' 在窗体模块顶部声明验证器
   Private validator As ClsFieldValidator
   
   ' 窗体加载时初始化
   Private Sub Form_Load()
       Set validator = New ClsFieldValidator
   End Sub
   
   ' 邮箱失去焦点时验证
   Private Sub txtEmail_LostFocus()
       Dim result As ValidationResult
       
       result = validator.Validate(Me.txtEmail, vtEmail)
       
       ' 更新状态显示
       UpdateValidationStatus Me.lblEmailStatus, result, validator.ErrorMessage
       HighlightTextBox Me.txtEmail, result
   End Sub
   
   ' 手机号失去焦点时验证
   Private Sub txtPhone_LostFocus()
       Dim result As ValidationResult
       
       result = validator.Validate(Me.txtPhone, vtMobile)
       
       UpdateValidationStatus Me.lblPhoneStatus, result, validator.ErrorMessage
       HighlightTextBox Me.txtPhone, result
   End Sub
   
   ' 提交前验证所有字段
   Private Sub btn提交_Click()
       Dim isValid As Boolean
       isValid = True
       
       ' 验证邮箱
       If validator.Validate(Me.txtEmail, vtEmail) <> vrValid Then
           isValid = False
           MsgBox "请输入有效的邮箱地址！", vbExclamation
           Me.txtEmail.SetFocus
           Exit Sub
       End If
       
       ' 验证手机号
       If validator.Validate(Me.txtPhone, vtMobile) <> vrValid Then
           isValid = False
           MsgBox "请输入有效的手机号码！", vbExclamation
           Me.txtPhone.SetFocus
           Exit Sub
       End If
       
       ' 所有验证通过
       If isValid Then
           ' 执行保存操作
           DoCmd.RunCommand acCmdSaveRecord
           MsgBox "保存成功！", vbInformation
       End If
   End Sub
   ```

---

### 🎯 教程 6：复杂验证规则

**场景**：密码强度验证（6-20位，必须包含字母和数字）

**步骤**：

1. **使用长度验证 + 自定义正则**
   ```vba
   Private Sub txtPassword_LostFocus()
       Dim result As ValidationResult
       
       ' 先验证长度
       validator.MinLength = 6
       validator.MaxLength = 20
       result = validator.Validate(Me.txtPassword, vtLength)
       
       If result = vrValid Then
           ' 再验证是否包含字母和数字
           validator.CustomPattern = "^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]{6,20}$"
           result = validator.Validate(Me.txtPassword, vtCustomRegex)
       End If
       
       ' 显示结果
       If result = vrValid Then
           UpdateValidationStatus Me.lblPwdStatus, vrValid
           HighlightTextBox Me.txtPassword, vrValid
       Else
           UpdateValidationStatus Me.lblPwdStatus, vrInvalid, "密码必须6-20位，包含字母和数字"
           HighlightTextBox Me.txtPassword, vrInvalid
       End If
   End Sub
   ```

---

## 高级技巧

### 🎯 教程 7：Azure 翻译集成

**场景**：为多语言应用添加实时翻译

**步骤**：

1. **配置 Azure 翻译服务**
   - 登录 [Azure Portal](https://portal.azure.com)
   - 创建"翻译器"资源
   - 获取密钥和端点

2. **配置模块**
   ```vba
   ' 在 modAzureTranslator 中修改
   Private Const AZURE_KEY As String = "你的密钥"
   Private Const AZURE_REGION As String = "eastus"
   Private Const AZURE_ENDPOINT As String = "你的端点"
   ```

3. **使用翻译功能**
   ```vba
   Private Sub btn翻译_Click()
       Dim sourceText As String
       Dim result As String
       
       sourceText = Me.txt原文
       
       ' 中文翻译为英文
       result = TranslateText(sourceText, "zh-Hans", "en")
       
       If Len(result) > 0 Then
           Me.txt译文 = result
       End If
   End Sub
   ```

4. **批量翻译示例**
   ```vba
   Private Sub btn批量翻译_Click()
       Dim rs As DAO.Recordset
       Dim translatedText As String
       
       Set rs = CurrentDb.OpenRecordset("SELECT * FROM 产品表")
       
       Do While Not rs.EOF
           ' 翻译产品名称
           translatedText = TranslateText(rs!产品名称, "zh-Hans", "en")
           
           ' 更新英文名称字段
           rs.Edit
           rs!英文名称 = translatedText
           rs.Update
           
           rs.MoveNext
       Loop
       
       rs.Close
       MsgBox "批量翻译完成！", vbInformation
   End Sub
   ```

---

### 🎯 教程 8：批量添加错误处理

**场景**：为现有项目的所有代码添加统一的错误处理

**步骤**：

1. **备份数据库**（重要！）
   ```
   复制当前数据库文件作为备份
   ```

2. **启用 VBA 项目访问**
   - 文件 → 选项 → 信任中心 → 信任中心设置
   - 宏设置 → 勾选"信任对 VBA 工程对象模型的访问"

3. **导入模块**
   ```
   导入 modVBETools.bas
   ```

4. **添加引用**
   ```
   Microsoft Visual Basic for Applications Extensibility 5.3
   ```

5. **执行批量添加**
   ```vba
   ' 在立即窗口或创建按钮执行
   AddErrorHandlingToAll
   ```

6. **检查结果**
   - 打开 VBA 编辑器
   - 检查各个过程
   - 确认错误处理代码已添加

**注意事项**：
- 此操作会修改代码，请务必备份
- 建议先在测试数据库上验证
- 已有错误处理的过程不会重复添加

---

### 🎯 教程 9：使用 ADO 事务处理

**场景**：订单系统，下单时需要同时更新库存

**步骤**：

1. **准备 SQL 语句**
   ```vba
   Private Sub btn下单_Click()
       Dim db As New ADOExecute
       Dim sql As String
       
       On Error GoTo ErrHandler
       
       ' 连接数据库
       db.Connect
       
       ' 插入订单（ADOExecute 自动处理事务）
       sql = "INSERT INTO 订单表 (客户ID, 产品ID, 数量) " & _
             "VALUES (" & Me.客户ID & ", " & Me.产品ID & ", " & Me.数量 & ")"
       db.ExecuteSQL sql
       
       ' 更新库存
       sql = "UPDATE 产品表 SET 库存 = 库存 - " & Me.数量 & _
             " WHERE 产品ID = " & Me.产品ID
       db.ExecuteSQL sql
       
       MsgBox "订单创建成功！", vbInformation
       Exit Sub
       
   ErrHandler:
       MsgBox "创建订单失败：" & Err.Description, vbCritical
   End Sub
   ```

**优点**：
- 自动事务管理
- 出错自动回滚
- 代码简洁

---

## 💡 最佳实践

### 通用建议

1. **错误处理**
   - 始终添加错误处理代码
   - 使用 modVBETools 批量添加

2. **代码注释**
   - 为关键代码添加注释
   - 说明参数含义和返回值

3. **测试**
   - 在测试环境中验证
   - 使用不同数据测试边界情况

4. **性能优化**
   - 大批量操作使用事务
   - 避免在循环中重复创建对象

### 模块特定建议

**basAutoNumStr**
- 在 BeforeInsert 事件中使用
- 考虑并发访问问题

**modExportToPPT**
- 控制查询返回的记录数（建议 < 100）
- 使用筛选器减少数据量

**ClsFieldValidator**
- 在 LostFocus 事件中验证
- 提交前再次验证所有字段

---

## 📚 更多资源

- [模块文档](Module-Documentation) - 查看所有模块详细说明
- [FAQ](FAQ) - 常见问题解答
- [GitHub Issues](https://github.com/miaowei2/accessdevelop/issues) - 提问和反馈

---

**有新的教程想法？** 欢迎提交 [Issue](https://github.com/miaowei2/accessdevelop/issues) 或 Pull Request！
