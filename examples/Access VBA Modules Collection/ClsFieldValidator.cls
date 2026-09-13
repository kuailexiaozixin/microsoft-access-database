VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFieldValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 类模块: ClsFieldValidator
Option Compare Database
Option Explicit

' 验证结果枚举
Public Enum ValidationResult
    vrValid = 0
    vrInvalid = 1
    vrEmpty = 2
End Enum

' 验证类型枚举
Public Enum validationType
    vtRequired = 1          ' 必填
    vtEmail = 2             ' 邮箱
    vtMobile = 3            ' 手机号
    vtIDCard = 4            ' 身份证号
    vtNumeric = 5           ' 纯数字
    vtLength = 6            ' 长度限制
    vtRange = 7             ' 数值范围
    vtCustomRegex = 8       ' 自定义正则
    vtDate = 9              ' 日期格式
End Enum

Private m_MinLength As Long
Private m_MaxLength As Long
Private m_MinValue As Double
Private m_MaxValue As Double
Private m_CustomPattern As String
Private m_ErrorMessage As String

' ========== 属性 ==========
Public Property Get errorMessage() As String
    errorMessage = m_ErrorMessage
End Property

Public Property Let MinLength(value As Long)
    m_MinLength = value
End Property

Public Property Let MaxLength(value As Long)
    m_MaxLength = value
End Property

Public Property Let MinValue(value As Double)
    m_MinValue = value
End Property

Public Property Let MaxValue(value As Double)
    m_MaxValue = value
End Property

Public Property Let CustomPattern(value As String)
    m_CustomPattern = value
End Property

' ========== 核心验证方法 ==========
Public Function Validate(ByVal inputValue As Variant, ByVal validationType As validationType) As ValidationResult
    Dim strValue As String
    strValue = Nz(inputValue, "")
    
    ' 清空上次错误信息
    m_ErrorMessage = ""
    
    Select Case validationType
        Case vtRequired
            Validate = ValidateRequired(strValue)
            
        Case vtEmail
            Validate = ValidateEmail(strValue)
            
        Case vtMobile
            Validate = ValidateMobile(strValue)
            
        Case vtIDCard
            Validate = ValidateIDCard(strValue)
            
        Case vtNumeric
            Validate = ValidateNumeric(strValue)
            
        Case vtLength
            Validate = ValidateLength(strValue)
            
        Case vtRange
            Validate = ValidateRange(strValue)
            
        Case vtCustomRegex
            Validate = ValidateRegex(strValue)
            
        Case vtDate
            Validate = ValidateDate(strValue)
            
        Case Else
            Validate = vrValid
    End Select
End Function

' ========== 具体验证规则 ==========

' 必填验证
Private Function ValidateRequired(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        m_ErrorMessage = "此字段为必填项"
        ValidateRequired = vrEmpty
    Else
        ValidateRequired = vrValid
    End If
End Function

' 邮箱验证
Private Function ValidateEmail(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateEmail = vrEmpty
        Exit Function
    End If
    
    ' 使用 VBScript.RegExp 进行正则验证
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    End With
    
    If regex.test(strValue) Then
        ValidateEmail = vrValid
    Else
        m_ErrorMessage = "请输入有效的邮箱地址"
        ValidateEmail = vrInvalid
    End If
    
    Set regex = Nothing
End Function

' 手机号验证 (中国大陆11位手机号)
Private Function ValidateMobile(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateMobile = vrEmpty
        Exit Function
    End If
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .Pattern = "^1[3-9]\d{9}$"
    End With
    
    If regex.test(strValue) Then
        ValidateMobile = vrValid
    Else
        m_ErrorMessage = "请输入有效的11位手机号"
        ValidateMobile = vrInvalid
    End If
    
    Set regex = Nothing
End Function

' 身份证号验证 (18位)
Private Function ValidateIDCard(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateIDCard = vrEmpty
        Exit Function
    End If
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        ' 18位身份证：6位地区码 + 8位生日 + 3位顺序码 + 1位校验码
        .Pattern = "^\d{6}(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{3}[\dXx]$"
    End With
    
    If regex.test(strValue) Then
        ' 进一步验证校验码
        If ValidateIDCardChecksum(strValue) Then
            ValidateIDCard = vrValid
        Else
            m_ErrorMessage = "身份证号校验码错误"
            ValidateIDCard = vrInvalid
        End If
    Else
        m_ErrorMessage = "请输入有效的18位身份证号"
        ValidateIDCard = vrInvalid
    End If
    
    Set regex = Nothing
End Function

' 身份证校验码算法
Private Function ValidateIDCardChecksum(strValue As String) As Boolean
    Dim weights As Variant
    Dim checkCodes As String
    Dim total As Long
    Dim i As Long
    
    weights = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
    checkCodes = "10X98765432"
    
    total = 0
    For i = 1 To 17
        total = total + CInt(Mid(strValue, i, 1)) * weights(i - 1)
    Next i
    
    Dim checkChar As String
    checkChar = Mid(checkCodes, (total Mod 11) + 1, 1)
    
    ValidateIDCardChecksum = (UCase(Mid(strValue, 18, 1)) = checkChar)
End Function

' 纯数字验证
Private Function ValidateNumeric(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateNumeric = vrEmpty
        Exit Function
    End If
    
    If IsNumeric(strValue) Then
        ValidateNumeric = vrValid
    Else
        m_ErrorMessage = "请输入有效的数字"
        ValidateNumeric = vrInvalid
    End If
End Function

' 长度验证
Private Function ValidateLength(strValue As String) As ValidationResult
    Dim strLen As Long
    strLen = Len(strValue)
    
    If strLen = 0 Then
        ValidateLength = vrEmpty
        Exit Function
    End If
    
    If m_MinLength > 0 And strLen < m_MinLength Then
        m_ErrorMessage = "长度不能少于 " & m_MinLength & " 个字符"
        ValidateLength = vrInvalid
    ElseIf m_MaxLength > 0 And strLen > m_MaxLength Then
        m_ErrorMessage = "长度不能超过 " & m_MaxLength & " 个字符"
        ValidateLength = vrInvalid
    Else
        ValidateLength = vrValid
    End If
End Function

' 数值范围验证
Private Function ValidateRange(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateRange = vrEmpty
        Exit Function
    End If
    
    If Not IsNumeric(strValue) Then
        m_ErrorMessage = "请输入有效的数字"
        ValidateRange = vrInvalid
        Exit Function
    End If
    
    Dim numValue As Double
    numValue = CDbl(strValue)
    
    If numValue < m_MinValue Then
        m_ErrorMessage = "数值不能小于 " & m_MinValue
        ValidateRange = vrInvalid
    ElseIf numValue > m_MaxValue Then
        m_ErrorMessage = "数值不能大于 " & m_MaxValue
        ValidateRange = vrInvalid
    Else
        ValidateRange = vrValid
    End If
End Function

' 自定义正则验证
Private Function ValidateRegex(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateRegex = vrEmpty
        Exit Function
    End If
    
    If Len(m_CustomPattern) = 0 Then
        ValidateRegex = vrValid
        Exit Function
    End If
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = m_CustomPattern
    End With
    
    If regex.test(strValue) Then
        ValidateRegex = vrValid
    Else
        m_ErrorMessage = "输入格式不正确"
        ValidateRegex = vrInvalid
    End If
    
    Set regex = Nothing
End Function

' 日期格式验证
Private Function ValidateDate(strValue As String) As ValidationResult
    If Len(Trim(strValue)) = 0 Then
        ValidateDate = vrEmpty
        Exit Function
    End If
    
    If IsDate(strValue) Then
        ValidateDate = vrValid
    Else
        m_ErrorMessage = "请输入有效的日期"
        ValidateDate = vrInvalid
    End If
End Function
