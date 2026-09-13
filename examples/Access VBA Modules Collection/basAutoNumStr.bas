Attribute VB_Name = "basAutoNumStr"
Option Compare Database
Option Explicit

Public Function AutoNumStr(TableName As String, _
        FieldName As String, _
        Digit As Integer, _
        Optional Prefixal As String, _
        Optional DateFormat As String)
    On Error GoTo ErrorHandler
    Dim strPrefixal As String
    Dim strTemp     As String
    
    strPrefixal = Prefixal
    If DateFormat <> "" Then strPrefixal = strPrefixal & Format(Date, DateFormat)
    If strPrefixal <> "" Then strTemp = "[" & FieldName & "] Like '" & strPrefixal & "*'"
    strTemp = Nz(DMax(FieldName, TableName, strTemp), "0")
    strTemp = Val(Mid(strTemp, Len(strPrefixal) + 1)) + 1
    strTemp = Format(strTemp, String(Digit, "0"))
    AutoNumStr = strPrefixal & strTemp
    
Done:
    Exit Function
    
ErrorHandler:
    ShowMsg "错误编号:  #" & Err & vbCrLf & _
        "错误来源:  " & Err.Source & vbCrLf & _
        "错误信息:  " & Err.Description, vbCritical, "出错"
    Resume Done
End Function


