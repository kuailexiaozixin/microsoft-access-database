' ============================================================================
' 调试日志模块模板（modDebugLog）—— Access 版「Debug.Print 三件套」
' ============================================================================
' 用途：自动化测试时的代码调试信息捕获。VBE 立即窗口内容无法被 COM 回读，
'       因此把调试输出写入日志文件（或被测库目录 .\vba_debug.log），Python 侧回读。
' 原则：Debug.Print > Debug.Assert > MsgBox；自动化场景禁止 MsgBox。
' 使用（代码化注入，勿手动粘贴）：
'   1. 本模块复制为 templates/modDebugLog.bas ；
'   2. 通过 scripts/rebuild_module_from_src.py 原子导入被测库（GBK 无 BOM + CRLF）；
'   3. 被测代码用 LogPrint / LogTimer / Watch 输出；结束时 LogClose；
'   4. Python 侧回读 .\vba_debug.log 做断言。
' 生产开关：#Const DEBUG_MODE = True 时输出，False 时空转（不写文件）。
' ============================================================================

Option Compare Database
Option Explicit

#Const DEBUG_MODE = True

Private pLogPath As String
Private pLogOpen As Boolean

' 初始化日志文件（默认 .\vba_debug.log）
Public Sub LogInit(Optional ByVal strFile As String = "vba_debug.log")
    On Error Resume Next
    pLogPath = CurrentProject.Path & "\" & strFile
    pLogOpen = False
    On Error GoTo 0
End Sub

' 输出一行（等价 Debug.Print 的文件版）
Public Sub LogPrint(ByVal strLine As String)
#If DEBUG_MODE Then
    LogEnsure
    On Error Resume Next
    Open pLogPath For Append As #1
    Print #1, Format$(Now, "hh:mm:ss") & " " & strLine
    Close #1
    On Error GoTo 0
#End If
End Sub

' 性能计时：开始
Public Function LogTimerStart() As Double
    LogTimerStart = Timer
End Function

' 性能计时：结束并输出耗时
Public Sub LogTimer(ByVal dblStart As Double, Optional ByVal strOp As String = "操作")
    LogPrint strOp & " 耗时: " & Format$(Timer - dblStart, "0.000") & " 秒"
End Sub

' 变量监视（等价 Watch）
Public Sub Watch(ByVal strName As String, ByVal vValue As Variant)
    LogPrint "[" & strName & "] = " & vValue
End Sub

' 错误捕获：On Error 分支里记录 Err 全貌（捕获代码报错信息）
Public Sub LogErr(ByVal strContext As String)
    LogPrint "错误@" & strContext & ": #" & Err.Number & " " & Err.Description & _
             " [源:" & Err.Source & " 行:" & Erl & "]"
End Sub

' 关闭日志（清空句柄）
Public Sub LogClose()
#If DEBUG_MODE Then
    On Error Resume Next
    Close #1
    pLogOpen = False
    On Error GoTo 0
#End If
End Sub

Private Sub LogEnsure()
    If Not pLogOpen Then
        On Error Resume Next
        Open pLogPath For Append As #1
        Close #1
        pLogOpen = True
        On Error GoTo 0
    End If
End Sub
