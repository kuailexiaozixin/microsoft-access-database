Attribute VB_Name = "modExportToExcel"
Option Compare Database
Option Explicit
Function ExportToExcel(ByVal WorkbookName As String, _
        Optional ByVal SheetName As String, _
        Optional ByVal FirstRange As String = "A1")
    On Error GoTo ErrorHandler
    Dim objApp As Object
    Dim objBook As Object
    Dim objSheet As Object
    Dim objRange As Object
    Dim strFileName As String
    Dim strExtName As String
    Dim lngRow As Long
    Dim lngColumn As Long
    Dim lngI As Long
    
    Const xlLastCell = 11
    Const xlCenter = -4108
    Const xlEdgeLeft = 7
    Const xlEdgeTop = 8
    Const xlEdgeBottom = 9
    Const xlEdgeRight = 10
    Const xlInsideVertical = 11
    Const xlInsideHorizontal = 12
    Const xlContinuous = 1
    Const xlDiagonalDown = 5
    Const xlDiagonalUp = 6
    Const xlNone = -4142
    
    
    strExtName = ".xls"
    If Val(Application.Version) > 11 Then strExtName = ".xlsx"
    
    With Application.FileDialog(2) 'msoFileDialogSaveAs
        .InitialFileName = WorkbookName & strExtName
        If Not .show Then
            Exit Function
        End If
        strFileName = .SelectedItems(1)
        If Not strFileName Like "*" & strExtName Then
            strFileName = strFileName & strExtName
        End If
        If Len(Dir(strFileName)) > 0 Then Kill strFileName
    End With
    
    '    MsgBox Screen.ActiveForm.Name
    If Forms.count > 0 Then
        DoCmd.SelectObject acForm, Forms(Forms.count - 1).name
    End If
    RunCommand acCmdSelectAllRecords
    RunCommand acCmdCopy
    DoEvents
    SendKeys "{TAB}", True
    Set objApp = CreateObject("Excel.Application")
    objApp.Visible = True
    Set objBook = objApp.Workbooks().Add()
    Set objSheet = objBook.Worksheets(1)
    objApp.DisplayAlerts = False
    Do Until objBook.Sheets.count = 1
        objBook.Sheets(objBook.Sheets.count).Delete
    Loop
    objApp.DisplayAlerts = True
    If Len(SheetName) = 0 Then
        SheetName = WorkbookName
    End If
    objSheet.name = SheetName
    objSheet.Range(FirstRange).Select
    '粘贴数据到excel
    objSheet.Paste
    objApp.ActiveCell.SpecialCells(xlLastCell).Select
    lngRow = objApp.ActiveCell.Row
    lngColumn = objApp.ActiveCell.Column
    '替换所有列标题中的" | "为换行符
    For lngI = 1 To lngColumn
        objSheet.cells(1, lngI) = Replace(objSheet.cells(1, lngI), " | ", Chr(10))
    Next
    '选中有数据区域
    Set objRange = objSheet.Range("A1", objApp.ActiveCell)
    objRange.Select
    '设置格式（行高，列高，对齐，框线）
    With objRange
        '        .RowHeight = 13.5
        .ColumnWidth = 50
        .EntireColumn.AutoFit
        '        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        '    .Interior.ColorIndex = -4142 'xlNone
    End With
    '第一行为列标题，行高加倍
    objSheet.rows(1).RowHeight = 27
    objApp.Range("A1").Select
    '冻结首行
    objApp.ActiveWindow.SplitRow = 1
    objApp.ActiveWindow.FreezePanes = True
    '
    objApp.Range(FirstRange).Select
    objBook.SaveAs strFileName
    
ExitHere:
    On Error Resume Next
    objApp.CopyMode = False
    Set objApp = Nothing
    Set objBook = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
    Resume ExitHere
End Function


