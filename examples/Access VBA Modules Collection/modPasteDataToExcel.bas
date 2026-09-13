Attribute VB_Name = "modPasteDataToExcel"
Option Compare Database
Option Explicit
Function PasteDataToExcel(ByVal WorkbookName As String, _
        Optional ByVal WorksheetName As String, _
        Optional ByVal ColumnHeaderBackColor As Variant, _
        Optional ByVal ColumnHeaderTwiceRowHeight As Boolean)
    On Error GoTo ErrorHandler
    Dim objApp          As Object
    Dim objBook         As Object
    Dim objSheet        As Object
    Dim objRange        As Object
    Dim strFileName     As String
    Dim strExtName      As String
    Dim lngRow          As Long
    Dim lngCol          As Long
    Dim lngI            As Long
    Dim varRowHeight    As Variant
    
    Const msoFileDialogSaveAs = 2
    Const xlThin = 2
    Const xlCenter = -4108
    Const xlNone = -4142
    Const xlContinuous = 1
    Const xlDiagonalDown = 5
    Const xlDiagonalUp = 6
    Const xlEdgeLeft = 7
    Const xlEdgeTop = 8
    Const xlEdgeBottom = 9
    Const xlEdgeRight = 10
    Const xlInsideVertical = 11
    Const xlInsideHorizontal = 12
    
    With Application.FileDialog(msoFileDialogSaveAs)
        .InitialFileName = WorkbookName
        If .show Then
            strFileName = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With
    
    DoCmd.Hourglass True
    
    RunCommand acCmdSelectAllRecords
    RunCommand acCmdCopy
    DoEvents
    SendKeys "{TAB}", True
    Set objApp = CreateObject("Excel.Application")
    objApp.Visible = True
    Set objBook = objApp.Workbooks().Add()
    Set objSheet = objBook.Worksheets(1)
    varRowHeight = objSheet.rows(1).RowHeight
    objApp.DisplayAlerts = False
    Do Until objBook.Sheets.count = 1
        objBook.Sheets(objBook.Sheets.count).Delete
    Loop
    objApp.DisplayAlerts = True
    If Len(WorksheetName) > 0 Then
        objSheet.name = WorksheetName
    Else
        objSheet.name = WorkbookName
    End If
    objSheet.Range("A1").Select
    objSheet.Paste
    lngRow = objSheet.UsedRange.rows.count
    lngCol = objSheet.UsedRange.Columns.count
    
    Set objRange = objSheet.Range("A1", objSheet.cells(lngRow, lngCol))
    With objRange
        .Select
        .RowHeight = varRowHeight
        .ColumnWidth = 100
        .EntireColumn.AutoFit
        '        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        '        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        '    .Interior.ColorIndex = -4142 'xlNone
    End With
    
    If ColumnHeaderTwiceRowHeight Then
        objSheet.rows(1).RowHeight = varRowHeight * 2
    End If
    If IsMissing(ColumnHeaderBackColor) Then
        ColumnHeaderBackColor = 15986395
    End If
    Set objRange = objSheet.cells(1, lngCol)
    objSheet.Range("A1", objRange).Interior.color = ColumnHeaderBackColor
    objApp.Range("A1").Select
    
    objApp.ActiveWindow.SplitRow = 1
    objApp.ActiveWindow.FreezePanes = True
    '
    objApp.Range("A1").Select
    objBook.SaveAs strFileName  ', xlOpenXMLTemplate
    '    objBook.Close
    '    objApp.Workbooks.Open strFileName
    
    
ExitHere:
    objApp.Visible = True
    DoCmd.Hourglass False
    Set objApp = Nothing
    Set objBook = Nothing
    Exit Function
    
ErrorHandler:
    If Err <> 1004 Then
        MsgBox Err.Description, vbCritical, "Error #" & Err
    End If
    Resume ExitHere
End Function


