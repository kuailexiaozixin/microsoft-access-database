Attribute VB_Name = "modHTMLExport"
Option Compare Database
Option Explicit

 

' 主导出函数
Public Sub ExportData(objectType As String, objectName As String, _
                      outputPath As String, Optional useTemplate As Boolean = False)
    
    If useTemplate Then
        ' 使用自定义模板
        Select Case objectType
            Case "Table"
                ExportTableCustom objectName, outputPath
            Case "Query"
                ExportQueryCustom objectName, outputPath
            Case Else
                MsgBox "不支持的对象类型！", vbExclamation
        End Select
    Else
        ' 使用内置方法
        Dim acType As AcOutputObjectType
        Select Case objectType
            Case "Table": acType = acOutputTable
            Case "Query": acType = acOutputQuery
            Case "Form": acType = acOutputForm
            Case "Report": acType = acOutputReport
            Case Else
                MsgBox "不支持的对象类型！", vbExclamation
                Exit Sub
        End Select
        
        DoCmd.OutputTo acType, objectName, acFormatHTML, outputPath
    End If
    
    MsgBox "导出完成：" & outputPath, vbInformation
End Sub

' 自定义表导出
Private Sub ExportTableCustom(tableName As String, outputPath As String)
    ' 实现细节参考方法2
End Sub

' 自定义查询导出
Private Sub ExportQueryCustom(QueryName As String, outputPath As String)
    ' 实现细节参考方法2
End Sub

