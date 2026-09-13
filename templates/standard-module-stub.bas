Attribute VB_Name = "basTemplate"
' ============================================================================
' 标准模块存根（vbext_ct_StdModule = type=1）
'
' 【存储】按 .src 约定保存为 UTF-8-BOM。
' 【导入】写进 Access 前必须转成 GBK 无 BOM + CRLF，二选一：
'         (a) Application.LoadFromText 5, "模块名", 文件   （覆盖式，推荐）
'         (b) VBComponents.Import(文件) + DoCmd.Save 5, "模块名"
'         带 BOM 直导会把首行变成 "锘緼ttribute" → 工程瘫痪、所有 Run 报"找不到过程"。
' 【原则】对外接口统一用 Function（不要用 Sub），避免 Application.Run 重复执行。
' 【原则】值可能是对象时，用 Sub + ByRef out，按 IsObject 选 Set / 等号，
'         禁止用等号直接接函数返回值（错 450）。
' ============================================================================

Option Explicit

' 示例：纯计算内核（无界面依赖，可被 TDD 覆盖）
Public Function AddInt(ByVal a As Long, ByVal b As Long) As Long
    ' 取整一律用 Int()：CLng/CInt 是银行家舍入，边界数据会进错箱且不报错。
    AddInt = a + b
End Function
