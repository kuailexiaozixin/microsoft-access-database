Attribute VB_Name = "modAzureTranslator"



'=============================================
' 模块：modAzureTranslator
' 功能：Azure翻译API调用
' 作者：Access开发
'=============================================
Option Compare Database
Option Explicit

'Azure配置（替换为你的密钥）
Private Const AZURE_KEY As String = "你的Azure翻译密钥"
Private Const AZURE_REGION As String = "eastus"  '你的区域
Private Const AZURE_ENDPOINT As String = "你的终结点 "

'翻译单条文本
Public Function TranslateText(sourceText As String, _
                              fromLang As String, _
                              toLang As String) As String
    On Error GoTo Err_Handler
    
    If Len(Trim(sourceText)) = 0 Then Exit Function
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    '构建API URL
    url = AZURE_ENDPOINT & "translate?api-version=3.0" & _
          "&from=" & fromLang & "&to=" & toLang
    
    '构建JSON请求体
    requestBody = "[{""Text"":""" & EscapeJson(sourceText) & """}]"
    
    '创建HTTP请求
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Ocp-Apim-Subscription-Key", AZURE_KEY
    http.setRequestHeader "Ocp-Apim-Subscription-Region", AZURE_REGION
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.send requestBody
    
    '解析响应
    If http.status = 200 Then
        response = http.responseText
        TranslateText = ParseTranslation(response)
    Else
        MsgBox "翻译失败: " & http.status & vbCrLf & http.responseText, vbCritical
        TranslateText = ""
    End If
    
    Set http = Nothing
    Exit Function
    
Err_Handler:
    MsgBox "翻译出错: " & Err.Description, vbCritical
    TranslateText = ""
End Function

'批量翻译资源表
Public Sub TranslateAllResources(toLang As String)
    Dim rs As DAO.Recordset
    Dim translated As String
    Dim count As Long
    Dim total As Long
    Dim strSQL As String
    Dim langField As String
    
    '确定目标语言字段
    Select Case toLang
        Case "en": langField = "EN"
        Case "ja": langField = "JA"
        Case "ko": langField = "KO"
        Case Else
            MsgBox "不支持的语言: " & toLang, vbExclamation
            Exit Sub
    End Select
    
    '查询未翻译的记录
    strSQL = "SELECT * FROM tblLanguage WHERE IsTranslated = False OR " & langField & " Is Null"
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rs.EOF Then
        MsgBox "没有需要翻译的内容", vbInformation
        rs.Close
        Exit Sub
    End If
    
    rs.MoveLast
    total = rs.RecordCount
    rs.MoveFirst
    
    '显示进度
    SysCmd acSysCmdInitMeter, "正在翻译...", total
    
    Do While Not rs.EOF
        count = count + 1
        SysCmd acSysCmdUpdateMeter, count
        
        '调用翻译
        translated = TranslateText(Nz(rs!cn, ""), "zh-Hans", toLang)
        
        If Len(translated) > 0 Then
            rs.Edit
            rs.Fields(langField).value = translated
            rs!IsTranslated = True
            rs!LastUpdate = Now()
            rs.Update
            
            Debug.Print count & "/" & total & ": " & rs!cn & " -> " & translated
        End If
        
        '避免频繁调用触发限流（每秒最多3次）
        If count Mod 3 = 0 Then
'            Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
        rs.MoveNext
        DoEvents
    Loop
    
    SysCmd acSysCmdRemoveMeter
    rs.Close
    
    MsgBox "翻译完成！共处理 " & count & " 条记录", vbInformation
End Sub

'批量翻译（带进度窗体版本）
Public Sub TranslateWithProgress(toLang As String)
    Dim rs As DAO.Recordset
    Dim translated As String
    Dim count As Long
    Dim total As Long
    Dim strSQL As String
    Dim langField As String
    
    '打开进度窗体（假设你已创建）
    DoCmd.OpenForm "frmProgress"
    
    Select Case toLang
        Case "en": langField = "EN"
        Case "ja": langField = "JA"
        Case "fr": langField = "FR"
        Case Else: Exit Sub
    End Select
    
    strSQL = "SELECT * FROM tblLanguage WHERE " & langField & " Is Null"
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not rs.EOF Then
        rs.MoveLast: total = rs.RecordCount: rs.MoveFirst
        
        Do While Not rs.EOF
            count = count + 1
            
            '更新进度条（假设你有ProgressUpdate函数）
'            ProgressUpdate Forms!frmProgress, count, total, _
                "翻译中: " & rs!TextKey
            
            translated = TranslateText(rs!cn, "zh-Hans", toLang)
            
            If Len(translated) > 0 Then
                rs.Edit
                rs.Fields(langField).value = translated
                rs!IsTranslated = True
                rs!LastUpdate = Now()
                rs.Update
            End If
            
            If count Mod 3 = 0 Then
'                Application.Wait (Now + TimeValue("0:00:01"))
            End If
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    DoCmd.Close acForm, "frmProgress"
    MsgBox "翻译完成！", vbInformation
End Sub

'转义JSON特殊字符
Private Function EscapeJson(text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    EscapeJson = text
End Function

'解析翻译结果JSON
Private Function ParseTranslation(json As String) As String
    Dim pos1 As Long, pos2 As Long
    
    '简单解析：查找 "text":"翻译结果"
    pos1 = InStr(1, json, """text"":""", vbTextCompare)
    If pos1 > 0 Then
        pos1 = pos1 + 8
        pos2 = InStr(pos1, json, """")
        If pos2 > pos1 Then
            ParseTranslation = Mid(json, pos1, pos2 - pos1)
            '还原转义字符
            ParseTranslation = Replace(ParseTranslation, "\n", vbCrLf)
            ParseTranslation = Replace(ParseTranslation, "\\", "\")
            ParseTranslation = Replace(ParseTranslation, "\""", """")
        End If
    End If
End Function

'测试翻译
Public Sub TestTranslate()
    Dim result As String
    result = TranslateText("客户管理系统", "zh-Hans", "FR")
    MsgBox "翻译结果: " & result, vbInformation
End Sub

