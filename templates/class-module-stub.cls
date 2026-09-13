VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' 类模块存根（vbext_ct_ClassModule = type=2）
'
' 【存储】按 .src 约定保存为 UTF-8-BOM（首 3 字节 EF BB BF）。
' 【导入】写进 Access 前必须转成 GBK 无 BOM + CRLF，再执行：
'         VBComponents.Remove(旧同名) → VBComponents.Import(新文件) → DoCmd.Save(5, 名)
'         一键脚本：python scripts/rebuild_module_from_src.py <主库.accdb> <本文件>
' 【为什么】文件带 BOM 导入时，首行 "VERSION 1.0 CLASS" 不被识别
'           → 类模块退化甚至丢失。旧资料说"Import 必然建成 type=1"，已实测推翻。
'           反证：clsTemplate.cls 写成 GBK 无 BOM 经 Import，得到 type=2 并重开复验通过。
' 【禁用】VBComponents.Add(2) 建类模块后改名 —— 改名抛 VBA 错误 53，不可用。
' 【禁用】手工剥前 9 行再 InsertLines(1, 正文) —— 大段代码会"部分写入"并在尾部留残片。
' 【命名·实测】VBComponents.Import 的组件名取自本文件头部的 Attribute VB_Name，
'         不是文件名（.cls / .bas 都一样：本模板文件名叫 class-module-stub.cls，
'         导入后组件名是 clsTemplate）。要改组件名就改这一行。
' 【比对】VBE 不保存 VERSION/BEGIN/END/Attribute/MultiUse 这些头部行，
'         所以"读回逐行比对"必须先剥掉它们，否则会误报"不一致"。
'         同理，重建脚本移除旧组件时要按 Attribute VB_Name 匹配，不能按文件名匹配。
' 【语法】续行符 _ 之后必须紧跟下一行，中间禁止插空行，
'         否则报语法错误并级联"在 End Sub/End Function/End 属性后只能出现注释"。
' ============================================================================

Option Explicit

Private Sub Class_Initialize()
    ' 初始化
End Sub

Private Sub Class_Terminate()
    ' 清理
End Sub

' 对外接口统一用 Function（不要用 Sub）：Application.Run 调 Sub 会触发重试、执行两遍。
Public Function Demo() As String
    Demo = "ok"
End Function
