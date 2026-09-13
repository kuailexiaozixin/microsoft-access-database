Attribute VB_Name = "modVBETools"
Option Compare Database
Option Explicit


' 引用必须勾选: Microsoft Visual Basic for Applications Extensibility 5.3

Private Const ERR_HANDLER_LABEL As String = "ErrorHandler:"
Private Const ERR_GOTO_STATEMENT As String = "On Error GoTo ErrorHandler"

' ==========================================
' 主入口：给当前数据库的所有模块添加错误处理
' ==========================================
Public Sub AddErrorHandlingToAll()
    Dim vbComp As VBComponent
    Dim count As Integer
    count = 0
    
    ' 遍历所有组件 (窗体, 报表, 模块, 类)
    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        ' 跳过没有代码的组件
        If vbComp.CodeModule.CountOfLines > 0 Then
            AddErrorHandlingToModule vbComp.CodeModule, count
        End If
    Next vbComp
    
    MsgBox "处理完成！共修改了 " & count & " 个过程。", vbInformation, "完成"
End Sub

' ==========================================
' 核心逻辑：给单个模块内的所有过程添加错误处理
' ==========================================
Private Sub AddErrorHandlingToModule(mdl As CodeModule, ByRef changeCount As Integer)
    Dim i As Long
    Dim procName As String
    Dim procKind As vbext_ProcKind
    Dim startLine As Long, bodyLine As Long, endLine As Long
    Dim strLine As String
    Dim hasErrorHandler As Boolean
    
    ' 从第一行遍历到最后一行
    i = 1
    Do While i < mdl.CountOfLines
        procName = mdl.ProcOfLine(i, procKind)
        
        ' 如果当前行属于某个过程
        If procName <> "" Then
            ' 获取过程的起始行、声明结束行(代码体开始)、总结束行
            startLine = mdl.ProcStartLine(procName, procKind)
            bodyLine = mdl.ProcBodyLine(procName, procKind)
            endLine = startLine + mdl.ProcCountLines(procName, procKind) - 1
            
            ' 1. 检查是否已经存在错误处理
            ' 我们在整个过程范围内搜索 "On Error GoTo" 或 "ErrorHandler:"
            hasErrorHandler = False
            If mdl.Find(ERR_GOTO_STATEMENT, startLine, 1, endLine, -1) Then hasErrorHandler = True
            If mdl.Find(ERR_HANDLER_LABEL, startLine, 1, endLine, -1) Then hasErrorHandler = True
            
            ' 2. 如果没有，则添加
            If Not hasErrorHandler Then
                ' 插入错误处理头部 (在声明行之后插入)
                ' bodyLine 是 "Sub xxx()" 的下一行，或者是注释后的第一行代码
                ' 我们通常希望插在 Sub 定义行的下一行
                
                ' 寻找真正的 Sub/Function 定义行之后
                Dim insertPos As Long
                insertPos = bodyLine
                
                ' 插入 On Error GoTo
                mdl.InsertLines insertPos, vbTab & ERR_GOTO_STATEMENT
                
                ' 插入错误处理尾部 (在 End Sub/End Function 之前)
                ' 注意：插入行后，行号会变化，所以要重新计算 endLine 或者倒序插入
                ' 这里简单处理：在 End Sub 前插入
                
                ' 重新获取结束行，因为刚才插入了一行
                endLine = startLine + mdl.ProcCountLines(procName, procKind) - 1
                
                Dim footerCode As String
                footerCode = ""
                footerCode = footerCode & vbTab & "Exit " & GetProcTypeString(mdl, procName, procKind) & vbCrLf
                footerCode = footerCode & ERR_HANDLER_LABEL & vbCrLf
                footerCode = footerCode & vbTab & "ShowMsg ""Error "" & Err.Number & "": "" & Err.Description & "" in " & procName & """, vbCritical" & vbCrLf
                footerCode = footerCode & vbTab & "Resume Next" ' 或者 Resume Exit_Handler
                
                ' 在 End Sub 之前插入
                mdl.InsertLines endLine, footerCode
                
                changeCount = changeCount + 1
            End If
            
            ' 跳过当前过程，继续下一个
            i = endLine + 1
        Else
            i = i + 1
        End If
    Loop
End Sub

' 辅助函数：判断是 Sub 还是 Function，用于生成 Exit Sub 或 Exit Function
Private Function GetProcTypeString(mdl As CodeModule, procName As String, procKind As vbext_ProcKind) As String
    Dim declLine As Long
    Dim declStr As String
    
    declLine = mdl.ProcBodyLine(procName, procKind)
    ' 获取定义行的内容 (注意：如果定义跨多行，这里只取了第一行，通常够用)
    ' 更严谨的做法是向前搜索直到找到 Sub/Function 关键字
    
    ' 简单判断：搜索该过程定义行是否包含 "Function"
    ' 注意：ProcBodyLine 返回的是过程体的第一行，我们需要找声明行
    ' 声明行通常在 BodyLine 之前，或者就是 BodyLine (如果没有注释)
    
    ' 这里简化处理：默认 Sub，如果包含 Function 则是 Function
    ' 实际上 VBE 没有直接属性告诉你是 Sub 还是 Function，需要解析文本
    
    Dim searchRange As Long
    searchRange = mdl.ProcCountLines(procName, procKind)
    Dim fullProcText As String
    ' 读取整个过程头部的几行来判断
    Dim startL As Long
    startL = mdl.ProcStartLine(procName, procKind)
    Dim headerText As String
    headerText = mdl.Lines(startL, mdl.ProcBodyLine(procName, procKind) - startL + 1)
    
    If InStr(1, headerText, "Function ", vbTextCompare) > 0 Then
        GetProcTypeString = "Function"
    ElseIf InStr(1, headerText, "Property ", vbTextCompare) > 0 Then
        GetProcTypeString = "Property"
    Else
        GetProcTypeString = "Sub"
    End If
End Function
