Attribute VB_Name = "basExportChart"
Option Compare Database
Option Explicit

Public Function ExportChartToExcel(controlNames As Variant, exportPath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim chartControl As Object
    Dim i As Integer
    
    ' 检查文件是否存在，如果存在则删除
    If Dir(exportPath) <> "" Then
        Kill exportPath
    End If
    
    ' 创建Excel应用程序实例
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' 如果需要，可以将其设置为False以隐藏Excel
    
    ' 创建新的工作簿
    Set excelWorkbook = excelApp.Workbooks.Add
    
    ' 遍历控件名称数组
    For i = LBound(controlNames) To UBound(controlNames)
        ' 获取当前控件
       ' Debug.Print controlNames(i)
        Set chartControl = Screen.ActiveForm.Controls(controlNames(i))
    
        
        If Not chartControl Is Nothing Then
            ' 创建新的工作表
            If i > 0 Then
                Set excelSheet = excelWorkbook.Sheets.Add(, excelWorkbook.Sheets(excelWorkbook.Sheets.Count))
            Else
                Set excelSheet = excelWorkbook.Sheets(1)
            End If
            
            ' 将图表控件复制到剪贴板
            chartControl.SetFocus
            DoCmd.RunCommand acCmdCopy
            
            ' 将剪贴板内容粘贴到Excel
            excelSheet.Paste
            excelSheet.Name = controlNames(i) ' 设置工作表名称为控件名称
            
        Else
             ExportChartToExcel = False
        End If
    Next i
    
    ' 保存Excel文件
    excelWorkbook.SaveAs exportPath
    excelWorkbook.Close
    excelApp.Quit
    ExportChartToExcel = True
ExitHere:
    ' 清理对象
    Set excelSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Function
ErrorHandler:
    MyMsgBox Err.Description, vbCritical, "Error"
    ExportChartToExcel = False
    Resume ExitHere
End Function
