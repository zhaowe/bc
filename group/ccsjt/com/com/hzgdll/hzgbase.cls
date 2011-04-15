VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hzgbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Conn As ADODB.Connection

'功能：根据DBClassType枚举类型判断数据源
'返回值：数据源连接字符串
'参数:DbClass(自定义的枚举值,传入参数)
Public Function GetConnStr(ByVal DbClass As Integer) As String
'Friend Function GetConnStr(ByVal DbClass As Integer) As String
    
    Select Case DbClass
    
    Case 41
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
    Case 46
        GetConnStr = "provider=sqloledb;server=10.254.0.46;database=mastersystem;uid=sa;pwd=szx6275;"
    Case 48
        GetConnStr = "provider=sqloledb;server=10.254.0.48;database=mastersystem;uid=sa;pwd=szx6275;"
    Case 91
        GetConnStr = "provider=sqloledb;server=10.254.0.91;database=mastersystem;uid=sa;pwd=szx6275;"
   
    Case Else
        Err.Raise 10001, "HzgDll.HzgBase.GetConnStr", "DBclass is undefined"
    End Select

End Function

'功能：打开database连接
'参数：strConn：ADO Connect 连接字符串(传入参数)，
Public Function DBConnect(ByVal strConn As String) As Long
'Private Function DBConnect(ByVal strConn As String) As Long
    Dim ErrorNo As Long
    ErrorNo = 0

    On Error GoTo ErrorHandler
    Set Conn = New ADODB.Connection
    Conn.Open strConn
    Conn.CursorLocation = adUseClient
    DBConnect = ErrorNo     'successful
    Exit Function
ErrorHandler:
    'Debug.Print Err.Number & ";" & Err.Source & ";" & Err.Description
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    If ErrorNo = 10009 Then
        ErrorNo = 10010
        Err.Raise ErrorNo, "HzgDll.HzgBase.DBConnect", "open database unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    DBConnect = ErrorNo     'unsuccessful
End Function

'功能：执行任何SQL操作
'返回值：查询结果对象
'参数:SqlSelect(Select语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeSQL(ByVal SqlStr As String, ByVal DbClass As Integer) As Variant
    
    Dim strConn As String
    Dim ErrorNo As Long
    Dim Rs As ADODB.Recordset
    
    ErrorNo = 0
    'get DB connect string
    strConn = GetConnStr(DbClass)
        
    On Error GoTo ErrorHandlerKnow
    'open database connect
    ErrorNo = DBConnect(strConn)
    On Error GoTo ErrorHandlerUnknow
    'execute SQL
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Set Rs = Conn.Execute(SqlStr)
    Set Rs.ActiveConnection = Nothing
    Conn.Close
    Set Conn = Nothing
    Set ExeSQL = Rs

    Exit Function
    
ErrorHandlerUnknow:
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If
    If Not Rs Is Nothing Then
        Set Rs = Nothing
    End If
    
    If ErrorNo = 10009 Then
        Err.Raise 10153, "hzgDll.HzgBase.ExeSQL", "excute SQL unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
End Function

'功能：工程内部调用:根据不同的数据库调用错误返回自定义的错误代码
'返回值：错误代码(ErrorNo)
'参数:iErrorNo(数据库系统产生的错误号,传入参数）
Public Function GetDBError(ByVal iErrorNo As Long) As Long
'Friend Function GetDBError(ByVal iErrorNo As Long) As Long
    Dim ErrorNo As Long
    
    'MsgBox Conn.Errors(0).Number & "===" & Conn.Errors(0).Description & _
            "#####" & Conn.Errors(0).NativeError & "$$$$$" & Conn.Errors(0).SQLState
            
    Select Case iErrorNo
        Case 11
        '数据库服务器找不到-genaral network error
            ErrorNo = 10002
            Err.Description = "genaral network error"
        Case 207
        '无效的字段名-Invalid Column Name
            ErrorNo = 10003
            Err.Description = "Invalid Column Name"
        Case 208
        '无效的表或视图名-Invalid Table,View Name
            ErrorNo = 10004
            Err.Description = "Invalid Table,View Name"
        Case 2627
        '插入纪录已经存在-can't insert duplicate key
            ErrorNo = 10005
            Err.Description = "Could not insert duplicate key"
        Case 2812
        '存储过程不存在-Could not find stored procedure
            ErrorNo = 10006
            Err.Description = "Could not find stored procedure"
        Case 4060
        '数据库名错误-Could not open database
            ErrorNo = 10007
            Err.Description = "Could not open database"
        Case 18456
        '数据库登录名或密码错误-login fail
            ErrorNo = 10008
            Err.Description = "Database login fail"
        Case 201
        '存储过程参数不全-Produce excepts parameter,which was not supplied
            ErrorNo = 10025
            Err.Description = "Produce excepts parameter,which was not supplied"
        Case 8114
        '用户参数与存储过程定义的参数类型不一致-Error Converting data type
            ErrorNo = 10026
            Err.Description = "Error Converting data type"
        Case 8162
        '用户参数定义的传递方向与存储过程参数定义的传递方向不一致
            ErrorNo = 10027
            Err.Description = "Error direction"
        Case 170
        'SQL语句语法错误-SQL statment is incorrect syntax
            ErrorNo = 10034
            Err.Description = "SQL statment is incorrect syntax"
        Case 8501
        'Database Server未启MSDTC服务
            ErrorNo = 10423
            Err.Description = "MSDTC on database server is unavailable"
        Case Else
        '其他错误-未知的数据库访问错误
            'ErrorNo = iErrorNo
            ErrorNo = 10009
            
    End Select
    GetDBError = ErrorNo
End Function




