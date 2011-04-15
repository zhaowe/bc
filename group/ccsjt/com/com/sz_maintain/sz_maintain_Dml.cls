VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sz_maintain_ClsDml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Conn As ADODB.Connection

'数据源枚举
Public Enum DBClassType
    XLS = 1
    hr = 2
    SZ_Maintain = 91
End Enum

'功能：根据DBClassType枚举类型判断数据源
'返回值：数据源连接字符串
'参数:DbClass(自定义的枚举值,传入参数)
Friend Function GetConnStr(ByVal DbClass As DBClassType) As String
    Select Case DbClass
    
    Case 1   '"内务管理"DB Server 上的XLS
        GetConnStr = ""
    Case 2   '"内务管理"DB Server 上的hr
        GetConnStr = ""
    Case 41   '"深圳维修管理"DB Server 上的SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=SZ_Maintain;uid=sa;pwd=;"
    Case 46   '"深圳维修管理"DB Server 上的SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.46;database=SZ_Maintain;uid=sa;pwd=;"
    Case 48   '"深圳维修管理"DB Server 上的SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.48;database=SZ_Maintain;uid=sa;pwd=;"
    Case 91   '"深圳维修管理"DB Server 上的SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.1.91;database=SZ_Maintain;uid=sa;pwd=;"

    Case Else
        Err.Raise 10001, "sz_maintain_Dml.sz_maintain_ClsDml.GetConnStr", "DBclass is undefined"
    End Select
End Function

'功能：查询纪录操作
'返回值：查询结果记录集
'参数:SqlSelect(Select语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeSelect(ByVal SqlSelect As String, Optional ByVal DbClass As DBClassType = 1) As Recordset
    
    Dim strConn As String
    Dim ErrorNo As Long
    Dim Rs As ADODB.Recordset
    
    ErrorNo = 0
    '检查是否Select语句
    If LCase(Left(Trim(SqlSelect), 7)) = "select " Then
        'get DB connect string
        strConn = GetConnStr(DbClass)
        
        On Error GoTo ErrorHandlerKnow
        'open database connect
        ErrorNo = DBConnect(strConn)
        On Error GoTo ErrorHandlerUnknow
        'execute SQL
        Set Rs = New ADODB.Recordset
        Rs.CursorLocation = adUseClient
        Set Rs = Conn.Execute(SqlSelect)
        Set Rs.ActiveConnection = Nothing
        Conn.Close
        Set Conn = Nothing
        Set ExeSelect = Rs
        Set Rs = Nothing
    Else
        'Sql语句不是select语句
        Err.Raise 10011, "sz_maintain_Dml.sz_maintain_ClsDml.ExeSelect", "SQL string isn't SELECT"
    End If
    Exit Function
    
ErrorHandlerUnknow:
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    'Debug.Print SqlSelect
    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If
    If Not Rs Is Nothing Then
        Set Rs = Nothing
    End If
    
    If ErrorNo = 10009 Then
        Err.Raise 10012, "sz_maintain_Dml.sz_maintain_ClsDml.ExeSelect", "excute SQL-select unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
End Function

'功能：插入纪录操作
'返回值：错误代码(ErrorNo)
'参数:SqlInsert(Insert语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeInsert(ByVal SqlInsert As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long
  
    
    ErrorNo = 0
    '检查是否Insert语句
    If LCase(Left(Trim(SqlInsert), 7)) = "insert " Then
        'get DB connect string
        strConn = GetConnStr(DbClass)
        
        On Error GoTo ErrorHandlerKnow
        ErrorNo = DBConnect(strConn)
        On Error GoTo ErrorHandlerUnknow
        'execute SQL
        Conn.Execute SqlInsert
        Conn.Close
        Set Conn = Nothing
        'GetObjectContext.SetComplete
    Else
        'sql语句不是insert语句
        ErrorNo = 10013
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeInsert", "SQL string isn't INSERT"
    End If
    'ExeInsert = ErrorNo
    Exit Function
    
ErrorHandlerUnknow:
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If
    'GetObjectContext.SetAbort
    If ErrorNo = 10009 Then
        ErrorNo = 10014
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeInsert", "excute SQL-insert unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    'ExeInsert = ErrorNo
End Function

'功能：修改纪录操作
'返回值：错误代码(ErrorNo)
'参数:SqlUpdate(Update语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeUpdate(ByVal SqlUpdate As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long

    ErrorNo = 0
    '检查是否update语句
    If LCase(Left(Trim(SqlUpdate), 7)) = "update " Then
        'get DB connect string
        strConn = GetConnStr(DbClass)
        
        On Error GoTo ErrorHandlerKnow
        ErrorNo = DBConnect(strConn)
        On Error GoTo ErrorHandlerUnknow
        'execute SQL
        Conn.Execute SqlUpdate
        Conn.Close
        Set Conn = Nothing
    Else
        'sql语句不是update语句
        ErrorNo = 10015
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeUpdate", "SQL string isn't UPDATE"
    End If
    'ExeUpdate = ErrorNo
    Exit Function
    
ErrorHandlerUnknow:
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If
    If ErrorNo = 10009 Then
        ErrorNo = 10016
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeUpdate", "excute SQL-update unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    'ExeUpdate = ErrorNo
End Function

'功能：删除纪录操作
'返回值：错误代码(ErrorNo)
'参数:SqlDelete(Delete语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeDelete(ByVal SqlDelete As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long
    
    ErrorNo = 0
    '检查是否delete语句
    If LCase(Left(Trim(SqlDelete), 7)) = "delete " Then
        'get DB connect string
        strConn = GetConnStr(DbClass)
        
        On Error GoTo ErrorHandlerKnow
        ErrorNo = DBConnect(strConn)
        On Error GoTo ErrorHandlerUnknow
        'execute SQL
        Conn.Execute SqlDelete
        Conn.Close
        Set Conn = Nothing
    Else
        'sql语句不是delete语句
        ErrorNo = 10017
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeDelete", "SQL string isn't DELETE"
    End If
    'ExeDelete = ErrorNo
    Exit Function
    
ErrorHandlerUnknow:
    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If
    If ErrorNo = 10009 Then
        ErrorNo = 10018
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.ExeDelete", "excute SQL-delete unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    'ExeDelete = ErrorNo
End Function

'功能：执行任何SQL操作
'返回值：查询结果对象
'参数:SqlSelect(Select语句，传入参数)；DbClass(自定义的枚举值,传入参数)
Public Function ExeSQL(ByVal SqlStr As String, Optional ByVal DbClass As DBClassType = 1) As Variant
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
        Err.Raise 10153, "sz_maintain_Dml.sz_maintain_ClsDml.ExeSQL", "excute SQL unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
End Function



'功能：获得Select语句的记录集,前提是先调用了ExeSelect方法
'返回值：记录集
'参数:无
'Public Property Get RsGet() As Recordset
'    Set RsGet = Rs
'End Property


'功能：工程内部调用:根据不同的数据库调用错误返回自定义的错误代码
'返回值：错误代码(ErrorNo)
'参数:iErrorNo(数据库系统产生的错误号,传入参数）
Friend Function GetDBError(ByVal iErrorNo As Long) As Long
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

'功能：打开database连接
'参数：strConn：ADO Connect 连接字符串(传入参数)，
Private Function DBConnect(ByVal strConn As String) As Long
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
        Err.Raise ErrorNo, "sz_maintain_Dml.sz_maintain_ClsDml.DBConnect", "open database unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    DBConnect = ErrorNo     'unsuccessful
End Function

'功能：取一个Guid
'参数：无
'返回值：Guid
Public Function GetGuid() As String
    Dim strSql As String
    Dim rsID As Recordset
    
    strSql = "declare @Guid uniqueidentifier select @Guid=newid() select @Guid"
    Set rsID = ExeSQL(strSql, 1)
    GetGuid = rsID(0)
End Function

'功能：取数据库时间
'参数：DBclass
'返回值：数据库时间
Public Function GetDBTime(Optional ByVal DbClass As DBClassType = 1) As Date
    Dim strSql As String
    Dim rsTime As Recordset
    
    strSql = "select GetDate()"
    Set rsTime = ExeSQL(strSql, DbClass)
    GetDBTime = rsTime(0)
End Function
'功能：执行任何SQL操作
'返回值：查询结果的记录集
'参数:SqlString(传入参数)；DbClass(自定义的枚举值,传入参数)
'Public Function SPRet(ByVal SqlStr As String, Optional ByVal DbClass As DBClassType = 1) As Recordset
'    Dim strConn As String
'    Dim ErrorNo As Long
'    Dim Rs As ADODB.Recordset
'
'    ErrorNo = 0
'    'get DB connect string
'    strConn = GetConnStr(DbClass)
'
'    On Error GoTo ErrorHandlerKnow
'    'open database connect
'    ErrorNo = DBConnect(strConn)
'    On Error GoTo ErrorHandlerUnknow
'    'execute SQL
'    Set Rs = New ADODB.Recordset
'    Rs.CursorLocation = adUseClient
'    Set Rs = Conn.Execute(SqlStr)
'    Set Rs.ActiveConnection = Nothing
'    Conn.Close
'    Set Conn = Nothing
'    Set SPRet = Rs
'    Set Rs = Nothing
'
'   Exit Function
    '
'ErrorHandlerUnknow:
'    ErrorNo = GetDBError(Conn.Errors(0).NativeError)
'    Err.Number = ErrorNo
'ErrorHandlerKnow:
'    ErrorNo = Err.Number
'    If Not Conn Is Nothing Then
'        Set Conn = Nothing
'    End If
'    If Not Rs Is Nothing Then
'        Set Rs = Nothing
'    End If
'
'    If ErrorNo = 10009 Then
'        Err.Raise 10153, "sz_maintain_Dml.sz_maintain_ClsDml.SPRet", "excute SQL unknow error"
'    Else
'        Err.Raise ErrorNo, Err.Source, Err.Description
'    End If
'End Function



