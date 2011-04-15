VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Conn As ADODB.Connection

'数据源枚举
Public Enum DBClassType
    NetBooking = 1      '网上订座
    ETicket = 2         '电子客票
    FlightDynamic = 3   '航班动态
    Price = 4           '票价
    Timetable = 5       '时刻表
    Payment = 6         '支付网关
    User = 7            '用户管理
    Error = 8           '错误信息
    Parameter = 9       '静态参数、业务参数
    InterTicket = 10    '国际机票
    Hotel = 11          '酒店
    Order = 12          '订单信息
    Tmp1 = 13           '库名未定
    Tmp2 = 14           '库名未定
    CAACGateway = 15    'CAAC Gateway
    Flight = 16         '机票预定信息
        
    XLS = 1
    hr = 2
    SZ_Maintain = 21
        
        
End Enum

'功能：根据DBClassType枚举类型判断数据源
'返回值：数据源连接字符串
'参数:DbClass(自定义的枚举值,传入参数)
Friend Function GetConnStr(ByVal DbClass As DBClassType) As String
    Select Case DbClass
    Case 1
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=NetBooking;uid=sa;pwd=;"
    Case 2
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Eticket;uid=sa;pwd=;"
    Case 3
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=FlightDynamic;uid=sa;pwd=;"
    Case 4
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Price;uid=sa;pwd=;"
    Case 5
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Timetable;uid=sa;pwd=;"
    Case 6
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Payment;uid=sa;pwd=;"
    Case 7
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=UserInfo;uid=sa;pwd=;"
    Case 8
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.34;database=ErrorManage;uid=sa;pwd=;"
    Case 9
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Parameter;uid=sa;pwd=;"
    Case 10
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=InterTicket;uid=sa;pwd=;"
    Case 11
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Hotel;uid=sa;pwd=;"
    Case 12
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=BizOrder;uid=sa;pwd=;"
    Case 13
        'GetConnStr = "provider=sqloledb;server=10.254.0.46;database=mastersystem;uid=sa;pwd=;"
        GetConnStr = "provider=sqloledb;server=127.0.0.1;database=Tmp1;uid=sa;pwd=szx6275;"
    Case 14
        'GetConnStr = "provider=sqloledb;server=10.254.0.46;database=mastersystem;uid=sa;pwd=;"
        GetConnStr = "provider=sqloledb;server=127.0.0.1;database=Tmp2;uid=sa;pwd=szx6275;"
    Case 15
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=CAACGateway;uid=sa;pwd=;"
    Case 16
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
        'GetConnStr = "provider=sqloledb;server=10.101.3.91;database=Flight;uid=sa;pwd=;"
    
    Case 20
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"
   
    Case 21   '"深圳维修管理"DB Server 上的SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=SZ_Maintain;uid=sa;pwd=szx6275;"


   
    Case Else
        Err.Raise 10001, "Com_DML.clsDML.GetConnStr", "DBclass is undefined"
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
        Set ExeSelect = Rs
        Set Rs = Nothing
        
    Else
        'Sql语句不是select语句
        Err.Raise 10011, "Com_DML.clsDML.ExeSelect", "SQL string isn't SELECT"
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
        Err.Raise 10012, "Com_DML.clsDML.ExeSelect", "excute SQL-select unknow error"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeInsert", "SQL string isn't INSERT"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeInsert", "excute SQL-insert unknow error"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeUpdate", "SQL string isn't UPDATE"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeUpdate", "excute SQL-update unknow error"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeDelete", "SQL string isn't DELETE"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.ExeDelete", "excute SQL-delete unknow error"
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
        Err.Raise 10153, "Com_DML.clsDML.ExeSQL", "excute SQL unknow error"
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
        Err.Raise ErrorNo, "Com_DML.clsDML.DBConnect", "open database unknow error"
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
    Set rsID = ExeSQL(strSql, 7)
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
'功能：将输入的内容字串变为格式化字串
'返回值：内容的格式化字串。
'参数:NeiRong(字符串，传入参数)
Public Function ExeGongaoXinwenFormat(ByVal NeiRong As String) As String
    
    Dim NeiRongStr As String
    Dim NeiRongStrSss As String
    Dim strnum, i As Integer
    
    NeiRongStr = NeiRong
    NeiRongStr = "&nbsp;&nbsp;&nbsp;&nbsp;" + Trim(NeiRongStr)
    'NeiRongStr = "&nbsp;" + Trim(NeiRongStr)
    strnum = Len(NeiRongStr)
    NeiRongStrSss = ""
    i = 1
    Do While i <= strnum And Len(NeiRongStr) > 0
        If Left(NeiRongStr, 1) <> Chr(13) Then
           NeiRongStrSss = NeiRongStrSss + Left(NeiRongStr, 1)
           NeiRongStr = Right(NeiRongStr, Len(NeiRongStr) - 1)
          Else
           NeiRongStrSss = NeiRongStrSss + "<BR>&nbsp;&nbsp;&nbsp;&nbsp;"
           NeiRongStr = Trim(Right(NeiRongStr, Len(NeiRongStr) - 1))
'           NeiRongStr = "&nbsp;&nbsp;&nbsp;&nbsp;" + Trim(NeiRongStr)
        End If
        i = i + 1
    Loop
    
    ExeGongaoXinwenFormat = NeiRongStrSss

End Function



