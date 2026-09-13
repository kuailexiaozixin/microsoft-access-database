Attribute VB_Name = "basSample"
Option Compare Database
Option Explicit

' VCS 通道回归测试夹具：提供 Hello() 供 vcs_call_vba 断言，
' 也是 ExportObject 错误 91 回归门禁的导出对象。
Public Function Hello() As String
    Hello = "VCS OK"
End Function
