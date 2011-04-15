VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsErrorManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ErrorNo As Long
Private ErrorSource As String
Private ErrorDescription As String
Private ObjData As Com_DML.clsDML
Private iReturn As Long
Const DbClass = 8

'ClassAType A���������
Public Enum ErrClassA
    Java = 1
    SQL = 2
    VB = 3
    VC = 4
    Vi = 5
    ���� = 6
End Enum

'�ڲ��ⲿ�������
Public Enum InOutType
    Inner = 1
    Outer = 2
End Enum

'���԰汾
Public Enum Language
    en = 1  'Ӣ��
    zh = 2  '��������
    zh_hk = 3 '��������
End Enum

'Э������
Public Enum ProtocolClass
    http = 1 'һ���htmlվ��
    wap = 2  'Wap�ֻ�վ��
End Enum

'�Զ������ö������
Public Enum Operate
    Delete = 1
    Insert = 2
    Restore = 3
    Update = 4
End Enum

'�����ѯ��Ҫ��
Public Enum QueryFlag
    Equal = 1 '����
    Exceed = 2  '����
    Include = 3  '����
    Less = 4 'С��
End Enum

'���ܣ���Error���������ӡ�ɾ�����޸ļ�¼����
'����ֵ:ִ�н��, 0:�ɹ�; ���򷵻ش������ ErrorNo
'���������User���û�Id   ErrNo�����������   ReasonIn���ڲ�����ԭ��
'          SolutionIn���ڲ�����취   ClassA��A��������
'          ClassB��B��������   PrgName�����������   ErrGoto������ʱָ���ļ�
'          ErrType���ڲ����ⲿ����   OperateType����������
'����������ErrNoBack�����Ӽ�¼ʱ�Զ�������ErrorNo
Public Function ErrorDeal(ByVal User As String, _
                          Optional ByVal ErrNo As Long = 0, _
                          Optional ByVal ReasonIn As String = "", _
                          Optional ByVal SolutionIn As String = "", _
                          Optional ByVal ClassA As ErrClassA = 2, _
                          Optional ByVal ClassB As String = "", _
                          Optional ByVal PrgName As String = "", _
                          Optional ByVal ErrGoto As String = "sorry.asp", _
                          Optional ByVal ErrType As InOutType = 1, _
                          Optional ByVal OperateType As Operate = 3, _
                          Optional ByRef ErrNoBack As Variant) As Long
    
    Dim SqlErrDeal As String
    Dim SqlLocale As String
    Dim LocalClassA As String
    Dim LocalErrType As String
    Dim ModifyTime As String
    Dim MaxErrNo As Long
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    Set ObjData = New Com_DML.clsDML
    
    'UserId������Ϊ��
    If User = "" Then
        ErrorNo = 10036  'UserIdΪ��
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "UserId is null"
        GoTo ErrValue
    End If
    
    '��ö������תΪ�ַ���
    Select Case ClassA
        Case 1
            LocalClassA = "java"
        Case 2
            LocalClassA = "sql"
        Case 3
            LocalClassA = "vb"
        Case 4
            LocalClassA = "vc"
        Case 5
            LocalClassA = "vi"
        Case 6
            LocalClassA = "other"
        Case Else
            ErrorNo = 10037 'A���Ͳ�ƥ��
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "ClassA Error"
            GoTo ErrValue
    End Select
    Select Case ErrType
    Case 1
        LocalErrType = "i"  '�ڲ�����
    Case 2
        LocalErrType = "o"  '�ⲿ����
    Case Else
        ErrorNo = 10038 '���ⲿ���Ͳ�ƥ��
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "ErrType Error"
        GoTo ErrValue
    End Select
    
    '��ModifyTimeΪDate����,Insert����а���"����"��"����"�ַ�������ֻ����λ,SQl������
    ModifyTime = Format(Now, "yyyy/mm/dd hh:mm:ss")  '�޸ļ�¼ʱ��
    
    '�жϲ�������
    Select Case OperateType
    Case 1 'Delete
'//***��ɾ����¼��ֻ�����Ӽ�¼��ɾ�����***//
        If Len(CStr(ErrNo)) >= 5 Then
            On Error GoTo Err_HanDleDll
            SqlLocale = "update localetype set DelTag='y' where errorno=" & ErrNo & ""
            iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)

            SqlErrDeal = "update error set DelTag='y' where errorno=" & ErrNo & ""
            iReturn = ObjData.ExeUpdate(SqlErrDeal, DbClass)
        Else
            ErrorNo = 10040 'ɾ��ErrorNO����
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "The Lenth of ErrorNo less than 5"
            GoTo ErrValue
        End If
'//**************************************//
    Case 2  'Insert
        '�ж���������Ƿ�Ϊ��
        iReturn = ErrorJudge(ReasonIn, SolutionIn, PrgName)
        If iReturn <> 0 Then
            ErrorNo = 10041  'ReasonIn��SolutionIn��PrgName��������п�ֵ
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "Private Function ErrorJudge Error"
            GoTo ErrValue
        End If
        On Error GoTo Err_HanDleDll
        MaxErrNo = GetErrorNo()
        ErrNoBack = MaxErrNo + 1
        
        If ErrGoto = "" Then
            ErrGoto = "sorry.asp"
        End If
        SqlErrDeal = "insert into Error (ErrorNo,ErrorReasonIn,ErrorSolutionIn, " & _
            "UserId,ClassAType,ClassBType,ErrorPrgName,LastModify,ErrorGoto,ErrorType)" & _
            " Values (" & ErrNoBack & ",'" & ReasonIn & "','" & SolutionIn & "','" & _
            User & "','" & LocalClassA & "','" & ClassB & "','" & PrgName & "','" & _
            ModifyTime & "','" & ErrGoto & "','" & LocalErrType & "')"
        
        iReturn = ObjData.ExeInsert(SqlErrDeal, DbClass)
        
    Case 3  'Restore
        On Error GoTo Err_HanDleDll
        SqlLocale = "update localetype set DelTag='n' where errorno=" & ErrNo & ""
        iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)
        
        SqlErrDeal = "update error set DelTag='n' where errorno=" & ErrNo & ""
        iReturn = ObjData.ExeUpdate(SqlErrDeal, DbClass)
    Case 4  'Update
        '�ж���������Ƿ�Ϊ��
        iReturn = ErrorJudge(ReasonIn, SolutionIn, PrgName)
        If iReturn <> 0 Then
            ErrorNo = 10041  'ReasonIn��SolutionIn��PrgName��������п�ֵ
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "Private Function ErrorJudge Error"
            GoTo ErrValue
        End If
'        If ErrNo = 0 Then
'            ErrorNo = 99990 'û�������¼����
'            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
'            ErrorDescription = "ErrorNo is null"
'            GoTo ErrValue
'        End If
        If ReasonIn <> "" Then
            SqlErrDeal = SqlErrDeal & "ErrorReasonIn='" & ReasonIn & "',"
        End If
        If SolutionIn <> "" Then
            SqlErrDeal = SqlErrDeal & "ErrorSolutionIn='" & SolutionIn & "',"
        End If
'�޸ļ�¼�ǲ���UserId
'        If User <> "" Then
'            SqlErrDeal = SqlErrDeal & "UserId='" & User & "',"
'        End If
        SqlErrDeal = SqlErrDeal & "ClassAType='" & LocalClassA & "',"
        If ClassB <> "" Then
            SqlErrDeal = SqlErrDeal & "ClassBType='" & ClassB & "',"
        End If
        If PrgName <> "" Then
            SqlErrDeal = SqlErrDeal & "ErrorPrgName='" & PrgName & "',"
        End If
        If ErrGoto <> "" Then
            SqlErrDeal = SqlErrDeal & "ErrorGoto='" & ErrGoto & "',"
        End If
        SqlErrDeal = SqlErrDeal & "ErrorType='" & LocalErrType & "',"
        SqlErrDeal = SqlErrDeal & "LastModify='" & ModifyTime & "'"
        SqlErrDeal = "update error set " & SqlErrDeal & "where ErrorNo=" & ErrNo & ""
        On Error GoTo Err_HanDleDll
        iReturn = ObjData.ExeUpdate(SqlErrDeal, DbClass)
    Case Else
        ErrorNo = 10043 '�޴˲�������
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "operateType error"
    End Select
    
    GoTo ErrValue
    
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10044 'ErrorDeal������������
    Err.Raise ErrorNo, "Com_ErrorManage.clsErrorManage.ErrorDeal", "other error"
    Exit Function
    
Err_HanDleDll:
    Set ObjData = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
ErrValue:
    Set ObjData = Nothing
    ErrorDeal = ErrorNo
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ���LocaleType���������ӡ�ɾ�����޸ļ�¼����
'����ֵ:ִ�н��, 0:�ɹ�; ���򷵻ش������ ErrorNo
'���������ErrNo�����������   Locale�����԰汾
'          Protocol��Э������   ErrOut���ⲿ��ʾ����
'          SolutionOut���ⲿ����취   OperateType����������
Public Function LocaleTypeDeal(ByVal ErrNo As Long, _
                               Optional ByVal Locale As Language = 2, _
                               Optional ByVal Protocol As ProtocolClass = 1, _
                               Optional ByVal ErrOut As String = "", _
                               Optional ByVal SolutionOut As String = "", _
                               Optional ByVal OperateType As Operate = 4) As Long
    
    Dim SqlLocale As String
    Dim LocalLanguage As String
    Dim LocalProtocol As String
    Dim ModifyTime As String
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    Set ObjData = New Com_DML.clsDML
        
    '��ö����ת��Ϊ�ַ���
    Select Case Locale
    Case 1  'Ӣ��
        LocalLanguage = "en"
    Case 2  '���ļ���
        LocalLanguage = "zh"
    Case 3  '���ķ���
        LocalLanguage = "zh-hk"
    End Select
    
    Select Case Protocol
    Case 1 'HTTP
        LocalProtocol = "http"
    Case 2 'WAP
        LocalProtocol = "wap"
    End Select
    
    ModifyTime = Format(Now, "yyyy/mm/dd hh:mm:ss")  '�޸ļ�¼ʱ��
    
    Select Case OperateType
    Case 1  'Delete
'//***ɾ����¼***//
        If Len(CStr(ErrNo)) >= 5 Then
            SqlLocale = "delete from LocaleType where ErrorNo=" & ErrNo & " and " & _
            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
            On Error GoTo Err_HanDleDll
            iReturn = ObjData.ExeDelete(SqlLocale, DbClass)
        Else
            ErrorNo = 10045 'ɾ��ErrorNO����
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
            ErrorDescription = "The Lenth of ErrorNo less than 5"
            GoTo ErrValue
        End If
'//***************//

'//***��ɾ����¼��ֻ�����Ӽ�¼��ɾ�����***//
'        If Len(CStr(ErrNo)) >= 5 Then
'            SqlLocale = "update LocaleType set DelTag='y' where ErrorNo=" & ErrNo & " and " & _
'            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
'            On Error GoTo Err_HanDleDll
'            iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)
'        Else
'            ErrorNo = 10045 'ɾ��ErrorNO����
'            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
'            ErrorDescription = "The Lenth of ErrorNo less than 5"
'            GoTo ErrValue
'        End If
'//**************************************//

    Case 2  'Insert
        '�ж���������Ƿ�Ϊ��
        iReturn = LocaleJudge(ErrOut, SolutionOut)
        If iReturn <> 0 Then
            ErrorNo = 10046  'ErrOut��SolutionOut�����������Ϊ��
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
            ErrorDescription = "Private Function ErrorJudge Error"
            GoTo ErrValue
        End If
        '**************************************
        '�ж�Ҫ�����ֵ�Ƿ���ڣ������������жϣ�
        Dim RsLocale As New ADODB.Recordset
        On Error GoTo Err_HanDleDll
        
        SqlLocale = "select * from LocaleType where ErrorNo=" & ErrNo & " and " & _
            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
            
        RsLocale.CursorLocation = adUseClient
        Set RsLocale = ObjData.ExeSelect(SqlLocale, DbClass)
        
        If Not RsLocale.EOF Then
            ErrorNo = 10047  'Ҫ����ļ�¼�Ѵ���
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleType"
            ErrorDescription = "insert duplicate value"
            RsLocale.Close
            Set RsLocale = Nothing
            GoTo ErrValue
        End If
        RsLocale.Close
        Set RsLocale = Nothing
        
        '�жϽ���
        '****************************
        SqlLocale = "insert into LocaleType (ErrorNo,LocaleType,ProtocolType," & _
            "ErrorNameOut,ErrorSolutionOut,LastModify)" & _
            " Values ('" & ErrNo & "','" & LocalLanguage & "','" & LocalProtocol & "','" & _
            ErrOut & "','" & SolutionOut & "','" & ModifyTime & "')"
        iReturn = ObjData.ExeInsert(SqlLocale, DbClass)
    Case 3  'Restore
        On Error GoTo Err_HanDleDll
        SqlLocale = "update LocaleType set DelTag='n' where ErrorNo=" & ErrNo & " and " & _
        "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
        iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)
    Case 4  'Update
        '�ж���������Ƿ�Ϊ��
        iReturn = LocaleJudge(ErrOut, SolutionOut)
        If iReturn <> 0 Then
            ErrorNo = 10046  'ErrOut��SolutionOut�����������Ϊ��
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleType"
            ErrorDescription = "Private Function LocaleJudge Error"
            GoTo ErrValue
        End If
        '************************************************
        ' ErrorNo��LocaleType��ProtocolType�ֶ�ֵ�����޸�
        '************************************************
        SqlLocale = "update LocaleType set "
        If ErrOut <> "" Then
            SqlLocale = SqlLocale & "ErrorNameOut='" & ErrOut & "',"
        End If
        If SolutionOut <> "" Then
            SqlLocale = SqlLocale & "ErrorSolutionOut='" & SolutionOut & "',"
        End If
        SqlLocale = SqlLocale & "LastModify='" & ModifyTime & "'"
        SqlLocale = SqlLocale & "where ErrorNo=" & ErrNo & "" & _
        " and LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
        On Error GoTo Err_HanDleDll
        iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)
    Case Else
        ErrorNo = 10048 '�޴˲�������
        ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
        ErrorDescription = "operateType error"
    End Select
    
    GoTo ErrValue
    
Err_Handle:
    ErrorNo = 10049 'LocaleTypeDeal����������
    Err.Raise ErrorNo, "Com_ErrorManage.clsErrorManage.LocaleTypeDeal", "LocaleTypeDeal Error"
    Exit Function
    
Err_HanDleDll:
    Set ObjData = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
ErrValue:
    Set ObjData = Nothing
    LocaleTypeDeal = ErrorNo
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��������������ѯ��������¼��
'����ֵ:ִ�н��, 0:�ɹ�; ���򷵻ش������ ErrorNo
'���������SelField��Ҫ��ѯ���ֶ�  SelFlag����ѯҪ��  SelValue����ѯֵ
'         DelSign��ɾ�����,="y"ʱ����Ѿ�ɾ����="n"ʱ���δɾ��
Public Function ErrorQuery(Optional ByVal SelField As String = "all", _
                           Optional ByVal SelFlag As QueryFlag = 3, _
                           Optional ByVal SelValue As String = "", _
                           Optional ByVal DelSign As String = "n") As Recordset
    Dim SqlQuery As String
    Dim StrFlag As String
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    Set ObjData = New Com_DML.clsDML
    
    If LCase(SelField) = "all" Or SelValue = "" Then
        'SqlQuery = "Select * From View_ErrorShow where DelTag ='" & DelSign & "'"
        SqlQuery = "Select * From View_ErrorShow where DelTag ='" & DelSign & "' order by ErrorNo Desc" '����������
    Else
        Select Case SelFlag
        Case 1 '=
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " = " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " = " & SelValue & " and DelTag ='" & DelSign & "' order by ErrorNo Desc" '����������
        Case 2 '>
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " > " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " > " & SelValue & " and DelTag ='" & DelSign & "'order by ErrorNo Desc" '����������
        Case 3 'like
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " like  '%" & SelValue & "%' and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " like  '%" & SelValue & "%' and DelTag ='" & DelSign & "'order by ErrorNo Desc" '����������
        Case 4 '<
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " < " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " < " & SelValue & " and DelTag ='" & DelSign & "'order by ErrorNo Desc" '����������
        End Select
    End If
    On Error GoTo Err_HanDleDll
    Set ErrorQuery = ObjData.ExeSelect(SqlQuery, DbClass)
    GoTo ErrValue
    
Err_Handle:
    ErrorNo = 10050 '�˴�������Ժ����¶���
    Err.Raise ErrorNo, "Com_ErrorManage.clsErrorManage.ErrorQuery", "ErrorQuery Error"
    Exit Function
    
Err_HanDleDll:
    Set ObjData = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ���ȡִ��ErrorDeal����Error����¼ʱ����ErrorNo
'����ֵ��ִ�н��, 0:�ɹ�; ���򷵻ش������ ErrorNo
'��������
Private Function GetErrorNo() As Long
    Dim Rst As ADODB.Recordset
    Dim MaxErrNo As Long
    Dim strSql As String
    Dim iTemp As Long
    
    On Error GoTo Err_HanDleDll
    strSql = "select ErrorNo from Error"
    Set Rst = ObjData.ExeSelect(strSql, DbClass)
    
    If Rst.BOF And Rst.EOF Then
        GetErrorNo = 10000
    Else
        strSql = "select max(ErrorNo) from Error"
        Set Rst = ObjData.ExeSelect(strSql, DbClass)
        GetErrorNo = Rst(0)
    End If
        
    Rst.Close
    Set Rst = Nothing
    Exit Function
    
Err_HanDleDll:
    ErrorNo = 10042 'δ��ȡ��ErrorNo�����Ӳ���ʧ��
    Err.Raise ErrorNo, "Com_ErrorManage.clsErrorManage.ErrorDeal", "private function GetErrorNo Error"
    'Err.Raise Err.Number, Err.Source, Err.Description
End Function

'���ܣ���ErrorDeal����Error����¼��������е�ReasonIn��SolutionIn��PrgName�ж��Ƿ�Ϊ��
'����ֵ:ִ�н��, 0:���������Ϊ��; ���򷵻ش������ ErrorNo
'���������ReasonIn����ӦReasonIn����   SolutionIn����ӦSolutionIn����  PrgName����ӦPrgName����
Private Function ErrorJudge(ByVal ReasonIn As String, ByVal SolutionIn As String, _
                            ByVal PrgName As String) As Long
    If ReasonIn = "" Or SolutionIn = "" Or PrgName = "" Then
        ErrorJudge = 1 '�����������Ϊ��
    Else
        ErrorJudge = 0
    End If
End Function

'���ܣ���LocaleTypeDeal����LocaleType����¼��������е�ErrOut��SolutionOut�ж��Ƿ�Ϊ��
'����ֵ:ִ�н��, 0:���������Ϊ��; ���򷵻ش������ ErrorNo
'���������ReasonOut����ӦErrOut����   SolutionOut����ӦSolutionOut����
Private Function LocaleJudge(ByVal ReasonOut As String, ByVal SolutionOut As String) As Long
    If ReasonOut = "" Or SolutionOut = "" Then
        LocaleJudge = 1 '�����������Ϊ��
    Else
        LocaleJudge = 0
    End If
End Function