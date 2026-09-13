Attribute VB_Name = "M_ValidationUI"

' 标准模块: M_ValidationUI
Option Compare Database
Option Explicit

' 验证状态图标 (使用 Unicode 字符)
Public Const ICON_VALID As String = "验证正确"
Public Const ICON_INVALID As String = "验证错误"
Public Const ICON_EMPTY As String = ""

' 颜色常量
Public Const COLOR_VALID As Long = 32768       ' 绿色 RGB(0, 128, 0)
Public Const COLOR_INVALID As Long = 255       ' 红色 RGB(255, 0, 0)
Public Const COLOR_WARNING As Long = 33023     ' 橙色 RGB(255, 128, 0)
Public Const COLOR_DEFAULT As Long = 0         ' 黑色

' 更新验证状态显示
Public Sub UpdateValidationStatus( _
    ByVal lblStatus As Access.Label, _
    ByVal result As ValidationResult, _
    Optional ByVal errorMessage As String = "")
    
    Select Case result
        Case vrValid
            With lblStatus
                .Caption = ICON_VALID
                .ForeColor = COLOR_VALID
                .ControlTipText = "验证通过"
            End With
            
        Case vrInvalid
            With lblStatus
                .Caption = ICON_INVALID
                .ForeColor = COLOR_INVALID
                .ControlTipText = IIf(Len(errorMessage) > 0, errorMessage, "验证失败")
            End With
            
        Case vrEmpty
            With lblStatus
                .Caption = ICON_EMPTY
                .ForeColor = COLOR_DEFAULT
                .ControlTipText = ""
            End With
    End Select
End Sub

' 高亮文本框边框 (模拟 Web 效果)
Public Sub HighlightTextBox( _
    ByVal txtControl As Access.TextBox, _
    ByVal result As ValidationResult)
    
    Select Case result
        Case vrValid
            txtControl.BorderColor = COLOR_VALID
            
        Case vrInvalid
            txtControl.BorderColor = COLOR_INVALID
            
        Case vrEmpty
            txtControl.BorderColor = COLOR_DEFAULT
    End Select
End Sub

' 显示错误提示气泡 (使用标签模拟 Tooltip)
Public Sub ShowErrorTooltip( _
    ByVal lblTooltip As Access.Label, _
    ByVal message As String, _
    ByVal show As Boolean)
    
    If show And Len(message) > 0 Then
        With lblTooltip
            .Caption = message
            .Visible = True
            .BackColor = RGB(255, 240, 240)  ' 浅红色背景
            .ForeColor = COLOR_INVALID
            .BorderColor = COLOR_INVALID
            .BorderStyle = 1  ' 实线边框
        End With
    Else
        lblTooltip.Visible = False
    End If
End Sub

' 验证整个表单，返回是否全部通过
Public Function ValidateForm(frm As Access.Form, ParamArray validations() As Variant) As Boolean
    Dim i As Long
    Dim allValid As Boolean
    Dim result As ValidationResult
    Dim validator As ClsFieldValidator
    
    allValid = True
    Set validator = New ClsFieldValidator
    
    ' validations 参数格式: txtControl, lblStatus, ValidationType, [可选参数...]
    ' 示例调用: ValidateForm(Me, Me.txtEmail, Me.lblEmailStatus, vtEmail, ...)
    
    For i = LBound(validations) To UBound(validations) Step 3
        Dim txtCtrl As Access.TextBox
        Dim lblCtrl As Access.Label
        Dim vType As validationType
        
        Set txtCtrl = validations(i)
        Set lblCtrl = validations(i + 1)
        vType = validations(i + 2)
        
        result = validator.Validate(txtCtrl.value, vType)
        UpdateValidationStatus lblCtrl, result, validator.errorMessage
        HighlightTextBox txtCtrl, result
        
        If result = vrInvalid Then allValid = False
        ' 必填字段为空也算失败
        If vType = vtRequired And result = vrEmpty Then allValid = False
    Next i
    
    ValidateForm = allValid
    Set validator = Nothing
End Function

