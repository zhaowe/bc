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

'���ܣ�����DBClassTypeö�������ж�����Դ
'����ֵ������Դ�����ַ���
'����:DbClass(�Զ����ö��ֵ,�������)
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

'���ܣ���database����
'������strConn��ADO Connect �����ַ���(�������)��
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

'���ܣ�ִ���κ�SQL����
'����ֵ����ѯ�������
'����:SqlSelect(Select��䣬�������)��DbClass(�Զ����ö��ֵ,�������)
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

'���ܣ������ڲ�����:���ݲ�ͬ�����ݿ���ô��󷵻��Զ���Ĵ������
'����ֵ���������(ErrorNo)
'����:iErrorNo(���ݿ�ϵͳ�����Ĵ����,���������
Public Function GetDBError(ByVal iErrorNo As Long) As Long
'Friend Function GetDBError(ByVal iErrorNo As Long) As Long
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



