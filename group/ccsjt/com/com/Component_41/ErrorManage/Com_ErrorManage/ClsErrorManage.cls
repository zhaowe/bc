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

'ClassAType A类错误类型
Public Enum ErrClassA
    Java = 1
    SQL = 2
    VB = 3
    VC = 4
    Vi = 5
    其它 = 6
End Enum

'内部外部错误分类
Public Enum InOutType
    Inner = 1
    Outer = 2
End Enum

'语言版本
Public Enum Language
    en = 1  '英文
    zh = 2  '简体中文
    zh_hk = 3 '繁体中文
End Enum

'协议类型
Public Enum ProtocolClass
    http = 1 '一般的html站点
    wap = 2  'Wap手机站点
End Enum

'自定义操作枚举类型
Public Enum Operate
    Delete = 1
    Insert = 2
    Restore = 3
    Update = 4
End Enum

'错误查询的要求
Public Enum QueryFlag
    Equal = 1 '等于
    Exceed = 2  '大于
    Include = 3  '包含
    Less = 4 '小于
End Enum

'功能：对Error表进行添加、删除和修改记录操作
'返回值:执行结果, 0:成功; 否则返回错误代码 ErrorNo
'传入参数：User：用户Id   ErrNo：错误码代号   ReasonIn：内部错误原因
'          SolutionIn：内部解决办法   ClassA：A错误类型
'          ClassB：B错误类型   PrgName：错误程序名   ErrGoto：出错时指向文件
'          ErrType：内部、外部类型   OperateType：操作类型
'传出参数：ErrNoBack：添加记录时自动产生的ErrorNo
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
    
    'UserId不允许为空
    If User = "" Then
        ErrorNo = 10036  'UserId为空
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "UserId is null"
        GoTo ErrValue
    End If
    
    '将枚举类型转为字符串
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
            ErrorNo = 10037 'A类型不匹配
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "ClassA Error"
            GoTo ErrValue
    End Select
    Select Case ErrType
    Case 1
        LocalErrType = "i"  '内部类型
    Case 2
        LocalErrType = "o"  '外部类型
    Case Else
        ErrorNo = 10038 '内外部类型不匹配
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "ErrType Error"
        GoTo ErrValue
    End Select
    
    '若ModifyTime为Date类型,Insert语句中包含"上午"或"下午"字符串且年只有两位,SQl语句出错
    ModifyTime = Format(Now, "yyyy/mm/dd hh:mm:ss")  '修改记录时间
    
    '判断操作类型
    Select Case OperateType
    Case 1 'Delete
'//***不删除记录，只是添加记录的删除标记***//
        If Len(CStr(ErrNo)) >= 5 Then
            On Error GoTo Err_HanDleDll
            SqlLocale = "update localetype set DelTag='y' where errorno=" & ErrNo & ""
            iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)

            SqlErrDeal = "update error set DelTag='y' where errorno=" & ErrNo & ""
            iReturn = ObjData.ExeUpdate(SqlErrDeal, DbClass)
        Else
            ErrorNo = 10040 '删除ErrorNO错误
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "The Lenth of ErrorNo less than 5"
            GoTo ErrValue
        End If
'//**************************************//
    Case 2  'Insert
        '判断输入参数是否为空
        iReturn = ErrorJudge(ReasonIn, SolutionIn, PrgName)
        If iReturn <> 0 Then
            ErrorNo = 10041  'ReasonIn、SolutionIn、PrgName输入参数有空值
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
        '判断输入参数是否为空
        iReturn = ErrorJudge(ReasonIn, SolutionIn, PrgName)
        If iReturn <> 0 Then
            ErrorNo = 10041  'ReasonIn、SolutionIn、PrgName输入参数有空值
            ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
            ErrorDescription = "Private Function ErrorJudge Error"
            GoTo ErrValue
        End If
'        If ErrNo = 0 Then
'            ErrorNo = 99990 '没有输入记录主键
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
'修改记录是不改UserId
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
        ErrorNo = 10043 '无此操作类型
        ErrorSource = "Com_ErrorManage.clsErrorManage.ErrorDeal"
        ErrorDescription = "operateType error"
    End Select
    
    GoTo ErrValue
    
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10044 'ErrorDeal方法其它错误
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

'功能：对LocaleType表进行添加、删除和修改记录操作
'返回值:执行结果, 0:成功; 否则返回错误代码 ErrorNo
'传入参数：ErrNo：错误码代号   Locale：语言版本
'          Protocol：协议类型   ErrOut：外部显示内容
'          SolutionOut：外部解决办法   OperateType：操作类型
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
        
    '将枚举型转化为字符串
    Select Case Locale
    Case 1  '英文
        LocalLanguage = "en"
    Case 2  '中文简体
        LocalLanguage = "zh"
    Case 3  '中文繁体
        LocalLanguage = "zh-hk"
    End Select
    
    Select Case Protocol
    Case 1 'HTTP
        LocalProtocol = "http"
    Case 2 'WAP
        LocalProtocol = "wap"
    End Select
    
    ModifyTime = Format(Now, "yyyy/mm/dd hh:mm:ss")  '修改记录时间
    
    Select Case OperateType
    Case 1  'Delete
'//***删除记录***//
        If Len(CStr(ErrNo)) >= 5 Then
            SqlLocale = "delete from LocaleType where ErrorNo=" & ErrNo & " and " & _
            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
            On Error GoTo Err_HanDleDll
            iReturn = ObjData.ExeDelete(SqlLocale, DbClass)
        Else
            ErrorNo = 10045 '删除ErrorNO错误
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
            ErrorDescription = "The Lenth of ErrorNo less than 5"
            GoTo ErrValue
        End If
'//***************//

'//***不删除记录，只是添加记录的删除标记***//
'        If Len(CStr(ErrNo)) >= 5 Then
'            SqlLocale = "update LocaleType set DelTag='y' where ErrorNo=" & ErrNo & " and " & _
'            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
'            On Error GoTo Err_HanDleDll
'            iReturn = ObjData.ExeUpdate(SqlLocale, DbClass)
'        Else
'            ErrorNo = 10045 '删除ErrorNO错误
'            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
'            ErrorDescription = "The Lenth of ErrorNo less than 5"
'            GoTo ErrValue
'        End If
'//**************************************//

    Case 2  'Insert
        '判断输入参数是否为空
        iReturn = LocaleJudge(ErrOut, SolutionOut)
        If iReturn <> 0 Then
            ErrorNo = 10046  'ErrOut、SolutionOut输入参数不能为空
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
            ErrorDescription = "Private Function ErrorJudge Error"
            GoTo ErrValue
        End If
        '**************************************
        '判断要插入的值是否存在（按三个主键判断）
        Dim RsLocale As New ADODB.Recordset
        On Error GoTo Err_HanDleDll
        
        SqlLocale = "select * from LocaleType where ErrorNo=" & ErrNo & " and " & _
            "LocaleType='" & LocalLanguage & "' and ProtocolType='" & LocalProtocol & "'"
            
        RsLocale.CursorLocation = adUseClient
        Set RsLocale = ObjData.ExeSelect(SqlLocale, DbClass)
        
        If Not RsLocale.EOF Then
            ErrorNo = 10047  '要插入的记录已存在
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleType"
            ErrorDescription = "insert duplicate value"
            RsLocale.Close
            Set RsLocale = Nothing
            GoTo ErrValue
        End If
        RsLocale.Close
        Set RsLocale = Nothing
        
        '判断结束
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
        '判断输入参数是否为空
        iReturn = LocaleJudge(ErrOut, SolutionOut)
        If iReturn <> 0 Then
            ErrorNo = 10046  'ErrOut、SolutionOut输入参数不能为空
            ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleType"
            ErrorDescription = "Private Function LocaleJudge Error"
            GoTo ErrValue
        End If
        '************************************************
        ' ErrorNo、LocaleType、ProtocolType字段值不可修改
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
        ErrorNo = 10048 '无此操作类型
        ErrorSource = "Com_ErrorManage.clsErrorManage.LocaleTypeDeal"
        ErrorDescription = "operateType error"
    End Select
    
    GoTo ErrValue
    
Err_Handle:
    ErrorNo = 10049 'LocaleTypeDeal中其它错误
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

'功能：根据输入参数查询错误代码记录集
'返回值:执行结果, 0:成功; 否则返回错误代码 ErrorNo
'传入参数：SelField：要查询的字段  SelFlag：查询要求  SelValue：查询值
'         DelSign：删除标记,="y"时标记已经删除，="n"时标记未删除
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
        SqlQuery = "Select * From View_ErrorShow where DelTag ='" & DelSign & "' order by ErrorNo Desc" '按降序排列
    Else
        Select Case SelFlag
        Case 1 '=
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " = " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " = " & SelValue & " and DelTag ='" & DelSign & "' order by ErrorNo Desc" '按降序排列
        Case 2 '>
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " > " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " > " & SelValue & " and DelTag ='" & DelSign & "'order by ErrorNo Desc" '按降序排列
        Case 3 'like
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " like  '%" & SelValue & "%' and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " like  '%" & SelValue & "%' and DelTag ='" & DelSign & "'order by ErrorNo Desc" '按降序排列
        Case 4 '<
            'SqlQuery = "Select * From View_ErrorShow where " & SelField & " < " & SelValue & " and DelTag ='" & DelSign & "'"
            SqlQuery = "Select * From View_ErrorShow where " & SelField & " < " & SelValue & " and DelTag ='" & DelSign & "'order by ErrorNo Desc" '按降序排列
        End Select
    End If
    On Error GoTo Err_HanDleDll
    Set ErrorQuery = ObjData.ExeSelect(SqlQuery, DbClass)
    GoTo ErrValue
    
Err_Handle:
    ErrorNo = 10050 '此错误代码以后将重新定义
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

'功能：获取执行ErrorDeal添加Error表纪录时产生ErrorNo
'返回值：执行结果, 0:成功; 否则返回错误代码 ErrorNo
'参数：无
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
    ErrorNo = 10042 '未能取得ErrorNo，添加操作失败
    Err.Raise ErrorNo, "Com_ErrorManage.clsErrorManage.ErrorDeal", "private function GetErrorNo Error"
    'Err.Raise Err.Number, Err.Source, Err.Description
End Function

'功能：对ErrorDeal添加Error表纪录输入参数中的ReasonIn、SolutionIn、PrgName判断是否为空
'返回值:执行结果, 0:输入参数不为空; 否则返回错误代码 ErrorNo
'输入参数：ReasonIn：对应ReasonIn参数   SolutionIn：对应SolutionIn参数  PrgName：对应PrgName参数
Private Function ErrorJudge(ByVal ReasonIn As String, ByVal SolutionIn As String, _
                            ByVal PrgName As String) As Long
    If ReasonIn = "" Or SolutionIn = "" Or PrgName = "" Then
        ErrorJudge = 1 '输入参数不能为空
    Else
        ErrorJudge = 0
    End If
End Function

'功能：对LocaleTypeDeal添加LocaleType表纪录输入参数中的ErrOut、SolutionOut判断是否为空
'返回值:执行结果, 0:输入参数不为空; 否则返回错误代码 ErrorNo
'输入参数：ReasonOut：对应ErrOut参数   SolutionOut：对应SolutionOut参数
Private Function LocaleJudge(ByVal ReasonOut As String, ByVal SolutionOut As String) As Long
    If ReasonOut = "" Or SolutionOut = "" Then
        LocaleJudge = 1 '输入参数不能为空
    Else
        LocaleJudge = 0
    End If
End Function
