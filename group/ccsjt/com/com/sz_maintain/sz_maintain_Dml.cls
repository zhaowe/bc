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

'����Դö��
Public Enum DBClassType
    XLS = 1
    hr = 2
    SZ_Maintain = 91
End Enum

'���ܣ�����DBClassTypeö�������ж�����Դ
'����ֵ������Դ�����ַ���
'����:DbClass(�Զ����ö��ֵ,�������)
Friend Function GetConnStr(ByVal DbClass As DBClassType) As String
    Select Case DbClass
    
    Case 1   '"�������"DB Server �ϵ�XLS
        GetConnStr = ""
    Case 2   '"�������"DB Server �ϵ�hr
        GetConnStr = ""
    Case 41   '"����ά�޹���"DB Server �ϵ�SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.41;database=SZ_Maintain;uid=sa;pwd=;"
    Case 46   '"����ά�޹���"DB Server �ϵ�SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.46;database=SZ_Maintain;uid=sa;pwd=;"
    Case 48   '"����ά�޹���"DB Server �ϵ�SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.0.48;database=SZ_Maintain;uid=sa;pwd=;"
    Case 91   '"����ά�޹���"DB Server �ϵ�SZ_Maintain
        GetConnStr = "provider=sqloledb;server=10.254.1.91;database=SZ_Maintain;uid=sa;pwd=;"

    Case Else
        Err.Raise 10001, "sz_maintain_Dml.sz_maintain_ClsDml.GetConnStr", "DBclass is undefined"
    End Select
End Function

'���ܣ���ѯ��¼����
'����ֵ����ѯ�����¼��
'����:SqlSelect(Select��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
Public Function ExeSelect(ByVal SqlSelect As String, Optional ByVal DbClass As DBClassType = 1) As Recordset
    
    Dim strConn As String
    Dim ErrorNo As Long
    Dim Rs As ADODB.Recordset
    
    ErrorNo = 0
    '����Ƿ�Select���
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
        'Sql��䲻��select���
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

'���ܣ������¼����
'����ֵ���������(ErrorNo)
'����:SqlInsert(Insert��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
Public Function ExeInsert(ByVal SqlInsert As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long
  
    
    ErrorNo = 0
    '����Ƿ�Insert���
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
        'sql��䲻��insert���
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

'���ܣ��޸ļ�¼����
'����ֵ���������(ErrorNo)
'����:SqlUpdate(Update��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
Public Function ExeUpdate(ByVal SqlUpdate As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long

    ErrorNo = 0
    '����Ƿ�update���
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
        'sql��䲻��update���
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

'���ܣ�ɾ����¼����
'����ֵ���������(ErrorNo)
'����:SqlDelete(Delete��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
Public Function ExeDelete(ByVal SqlDelete As String, Optional ByVal DbClass As DBClassType = 1) As Long
    Dim strConn As String
    Dim ErrorNo As Long
    
    ErrorNo = 0
    '����Ƿ�delete���
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
        'sql��䲻��delete���
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

'���ܣ�ִ���κ�SQL����
'����ֵ����ѯ�������
'����:SqlSelect(Select��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
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



'���ܣ����Select���ļ�¼��,ǰ�����ȵ�����ExeSelect����
'����ֵ����¼��
'����:��
'Public Property Get RsGet() As Recordset
'    Set RsGet = Rs
'End Property


'���ܣ������ڲ�����:���ݲ�ͬ�����ݿ���ô��󷵻��Զ���Ĵ������
'����ֵ���������(ErrorNo)
'����:iErrorNo(���ݿ�ϵͳ�����Ĵ����,���������
Friend Function GetDBError(ByVal iErrorNo As Long) As Long
    Dim ErrorNo As Long
    
    'MsgBox Conn.Errors(0).Number & "===" & Conn.Errors(0).Description & _
            "#####" & Conn.Errors(0).NativeError & "$$$$$" & Conn.Errors(0).SQLState
            
    Select Case iErrorNo
        Case 11
        '���ݿ�������Ҳ���-genaral network error
            ErrorNo = 10002
            Err.Description = "genaral network error"
        Case 207
        '��Ч���ֶ���-Invalid Column Name
            ErrorNo = 10003
            Err.Description = "Invalid Column Name"
        Case 208
        '��Ч�ı�����ͼ��-Invalid Table,View Name
            ErrorNo = 10004
            Err.Description = "Invalid Table,View Name"
        Case 2627
        '�����¼�Ѿ�����-can't insert duplicate key
            ErrorNo = 10005
            Err.Description = "Could not insert duplicate key"
        Case 2812
        '�洢���̲�����-Could not find stored procedure
            ErrorNo = 10006
            Err.Description = "Could not find stored procedure"
        Case 4060
        '���ݿ�������-Could not open database
            ErrorNo = 10007
            Err.Description = "Could not open database"
        Case 18456
        '���ݿ��¼�����������-login fail
            ErrorNo = 10008
            Err.Description = "Database login fail"
        Case 201
        '�洢���̲�����ȫ-Produce excepts parameter,which was not supplied
            ErrorNo = 10025
            Err.Description = "Produce excepts parameter,which was not supplied"
        Case 8114
        '�û�������洢���̶���Ĳ������Ͳ�һ��-Error Converting data type
            ErrorNo = 10026
            Err.Description = "Error Converting data type"
        Case 8162
        '�û���������Ĵ��ݷ�����洢���̲�������Ĵ��ݷ���һ��
            ErrorNo = 10027
            Err.Description = "Error direction"
        Case 170
        'SQL����﷨����-SQL statment is incorrect syntax
            ErrorNo = 10034
            Err.Description = "SQL statment is incorrect syntax"
        Case 8501
        'Database Serverδ��MSDTC����
            ErrorNo = 10423
            Err.Description = "MSDTC on database server is unavailable"
        Case Else
        '��������-δ֪�����ݿ���ʴ���
            'ErrorNo = iErrorNo
            ErrorNo = 10009
            
    End Select
    GetDBError = ErrorNo
End Function

'���ܣ���database����
'������strConn��ADO Connect �����ַ���(�������)��
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

'���ܣ�ȡһ��Guid
'��������
'����ֵ��Guid
Public Function GetGuid() As String
    Dim strSql As String
    Dim rsID As Recordset
    
    strSql = "declare @Guid uniqueidentifier select @Guid=newid() select @Guid"
    Set rsID = ExeSQL(strSql, 1)
    GetGuid = rsID(0)
End Function

'���ܣ�ȡ���ݿ�ʱ��
'������DBclass
'����ֵ�����ݿ�ʱ��
Public Function GetDBTime(Optional ByVal DbClass As DBClassType = 1) As Date
    Dim strSql As String
    Dim rsTime As Recordset
    
    strSql = "select GetDate()"
    Set rsTime = ExeSQL(strSql, DbClass)
    GetDBTime = rsTime(0)
End Function
'���ܣ�ִ���κ�SQL����
'����ֵ����ѯ����ļ�¼��
'����:SqlString(�������)��DbClass(�Զ����ö��ֵ,�������)
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


