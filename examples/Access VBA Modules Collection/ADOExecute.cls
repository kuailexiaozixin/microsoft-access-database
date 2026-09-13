VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

 

' 定义属性
Private m_Connection As Object
Private m_ConnectionString As String

' 定义方法
Public Sub Connect()
    ' 创建连接对象
    Set m_Connection = CreateObject("ADODB.Connection")
    
    ' 连接到数据库
    m_Connection.Open m_ConnectionString
End Sub

Public Sub ExecuteSQL(sql As String)
    Dim cmd As Object
    Dim param As Object
    Dim i As Integer
    
    ' 创建命令对象
    Set cmd = CreateObject("ADODB.Command")
    
    ' 设置连接对象和 SQL 语句
    cmd.ActiveConnection = m_Connection
    cmd.CommandText = sql

    
    ' 开始事务
    m_Connection.BeginTrans
    
    On Error GoTo RollbackTransaction
    
    ' 执行 SQL 语句
    cmd.Execute
    
    ' 提交事务
    m_Connection.CommitTrans
    
    Exit Sub
    
RollbackTransaction:
    ' 回滚事务
    m_Connection.RollbackTrans
    Err.Raise vbObjectError + 1, "Database.Execute", Err.Description
End Sub
Public Function GetValues(sql As String) As Variant
On Error GoTo ErrorHandler
    Dim rst As Object
    Dim values As Variant
    
      ' 创建命令对象
    Set rst = CreateObject("ADODB.Recordset")
    
'    ' 设置连接对象和 SQL 语句
'    rst.ActiveConnection = m_Connection
    
    ' 执行 SQL 查询
    rst.Open sql, m_Connection
    
    ' 检查是否有记录
    If Not rst.EOF Then
        values = rst.Fields(0).Value
    End If
    
    ' 关闭记录集
   rst.Close
    
    ' 返回值数组
    GetValues = values
ExitHere:
    
    Exit Function
ErrorHandler:
    MyMsgBox Err.Description, vbCritical, "#Error"
    Resume ExitHere
End Function

' 定义构造函数
Public Sub Class_Initialize()
    ' 设置默认的连接字符串
    m_ConnectionString = "Provider=SQLOLEDB" & _
        ";Data Source=" & GetProperty("USysFrmSetup", "DbServer", "") & _
        ";Initial Catalog=" & GetProperty("USysFrmSetup", "DbDatabase", "") & _
        ";User ID=" & GetProperty("USysFrmSetup", "DbUserName", "") & _
        ";Password=" & GetProperty("USysFrmSetup", "DbPassword", "")
End Sub
Public Sub Disconnect()
    ' 关闭连接
    m_Connection.Close
    
    ' 释放对象
    Set m_Connection = Nothing
End Sub

' 设置连接字符串
Public Property Let connectionString(connectionString As String)
    m_ConnectionString = connectionString
End Property
 
