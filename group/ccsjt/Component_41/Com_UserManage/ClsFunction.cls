VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ClsFunction1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************
'* Author:LiangYe And Mr.wei     *
'* Modify Date:2000-10-30        *
'* Purpose:AMS用户管理            *
'*********************************
Option Explicit

Private Const DbClass = 7                       '数据库指向UserManage
Private ErrorNo As Long
Private Const AdminFunction = "UserManage"      '用户管理功能的描述

'功能:根据条件查询视图View_FunctionInfo信息记录集
'返回值:功能信息记录集
'输入参数：TextFileld:字段  TiaoJian:条件值
Public Function SearchFunctionInfo(ByVal TextField As String, ByVal TiaoJian As String) As Recordset
    Dim ObjDml As Object
    Dim strSql As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    If TextField = "" Or TiaoJian = "" Then
        strSql = "select * from View_functionInfo order by OrderNum"
    Else
        TextField = "%" & TextField & "%"
        strSql = "select * from View_functionInfo where " & TiaoJian & " like '" & TextField & "'" _
                        & " order by OrderNum"
    End If
    On Error GoTo ErrorHandlerKnow
    Set SearchFunctionInfo = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchFunctionInfo.EOF Then
        ErrorNo = 10791
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据Locale获取指定视图View_FunctionInfo信息记录集
'返回值:功能信息记录集
'输入参数：Locale
Public Function GetAllFunction(ByVal Locale As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If Locale = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "select * from View_FunctionInfo where locale='" & Locale & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetAllFunction = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据FunctionID获取指定视图View_FunctionInfo信息记录集
'返回值:信息记录集
'输入参数：FunctionID,Locale
Public Function GetFunctionInfo(ByVal FunctionID As String, Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If FunctionID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    If Locale = "" Then
        strSql = "select * from View_FunctionInfo where FunctionID='" & FunctionID & "'" _
                    & " order by orderNum"
    Else
        strSql = "select * from View_FunctionInfo where FunctionID='" & FunctionID & "'" _
                    & " and locale='" & Locale & "'" & " order by orderNum"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetFunctionInfo = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetFunctionInfo.EOF Then
        ErrorNo = 10791
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据FunctionID获取FunctionLocale信息记录集
'返回值:FunctionLocale信息记录集
'输入参数：FunctionID
Public Function GetFunctionLocale(ByVal FunctionID As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If FunctionID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "select * from FunctionLocale where FunctionID='" & FunctionID & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetFunctionLocale = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据FunctionID,Locale获取指定功能名FunctionName
'返回值:指定语言功能名
'输入参数：FunctionID,Locale
Public Function GetFunctionName(ByVal FunctionID As String, ByVal Locale As String) As String
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    If FunctionID = "" Or Locale = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "select * from FunctionLocale where FunctionID='" & FunctionID & "'" _
                & " and Locale='" & Locale & "'"
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    GetFunctionName = Rs("FunctionName")
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not Rs Is Nothing Then
        Set Rs = Nothing
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据FunctionID获取指定视图信息记录集
'返回值:信息记录集
'输入参数：FunctionID,UseObject,locale
Public Function GetFunctionUser(ByVal FunctionID As String, ByVal UseObject As String, _
                                Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If FunctionID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    If Locale = "" Then
        strSql = "select * from View_FunctionUser where FunctionID='" & FunctionID & "'" _
                & " and UseObject='" & UseObject & "'"
    Else
        strSql = "select * from View_FunctionUser where FunctionID='" & FunctionID & "'" _
                & " and locale='" & Locale & "' and UseObject='" & UseObject & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetFunctionUser = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能:根据FunctionID获取指定视图View_FunctionGroup信息记录集
'返回值:信息记录集
'输入参数：FunctionID,Locale
Public Function GetFunctionGroup(ByVal FunctionID As String, Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If FunctionID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    If Locale = "" Then
        strSql = "select * from View_FunctionGroup  where FunctionID='" & FunctionID & "'"
    Else
        strSql = "select * from View_FunctionGroup  where FunctionID='" & FunctionID & "'" _
                & " and locale='" & Locale & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetFunctionGroup = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：添加新纪录到Function表
'返回值：FunctionID
'输入参数:Description
Public Function AddFunction(ByVal Description As String, ByVal FunctionType As String, Optional ByVal Conflict As Integer) As String
    Dim FunctionID As String
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim OrderNum As Integer
    Dim i As Integer
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    Description = FiltQuoteChar(Description)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    
    On Error GoTo ErrorHandlerKnow
    FunctionID = ObjDml.GetGuid()
    OrderNum = GetMaxNum()
    strSql = "Insert into [Function] (FunctionID,Description,OrderNum,Conflict,FunctionType) values " _
            & "('" & FunctionID & "','" & Description & "'," & OrderNum & "," & Conflict & ",'" & FunctionType & "')"
    '执行插入Function表操作
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    AddFunction = FunctionID
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：添加新纪录到Function、FunctionLocale和FunctionRelation表
'返回值：FunctionId
'输入参数:FunctionName功能名
'           Description描述
'           fFunctionID父功能ID
'           Conflict功能冲突
'           FunctionType功能类型
'           Locale语言版本
Public Function AddFunctionAll(ByVal FunctionName As String, _
                            ByVal Description As String, ByVal fFunctionID As String, _
                            Optional ByVal Conflict As Integer, _
                            Optional ByVal FunctionType As String, _
                            Optional ByVal Locale As String = "zh") As String
    Dim FunctionID As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim Str As String
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    FunctionName = FiltQuoteChar(FunctionName)
    Description = FiltQuoteChar(Description)
    FunctionID = AddFunction(Description, FunctionType, Conflict)
    Ret = AddFunctionLocale(FunctionID, FunctionName)
    Str = AddRelation(FunctionID, fFunctionID)
    AddFunctionAll = FunctionID
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    GetObjectContext.SetAbort
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：添加FunctionLocale表
'返回值：无
'输入参数:
'       FunctionID
'       FunctionName
'       Locale
Public Function AddFunctionLocale(ByVal FunctionID As String, ByVal FunctionName As String, _
                                Optional ByVal Locale As String = "zh")
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    FunctionName = FiltQuoteChar(FunctionName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "Insert into FunctionLocale (FunctionID,FunctionName,Locale) values " _
            & "('" & FunctionID & "','" & FunctionName & "','" & Locale & "')"
    '执行插入FunctionLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：修改Function表信息
'返回值：无
'输入参数:FunctionID, FunctionType
Public Function EditFunction(ByVal FunctionID As String, ByVal FunctionType As String)
    Dim ObjDml As Object
    Dim strSql, strSqlA As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim ObjRs As Recordset
    Set ObjRs = CreateObject("Adodb.Recordset")
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
       
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "Update [Function] set FunctionType='" & FunctionType & "'" _
            & " where FunctionId='" & FunctionID & "'"
    strSqlA = "select FunctionType from [Function] where FunctionID='" & FunctionID & "'"
    '执行修改FunctionLocale表操作
    On Error GoTo ErrorHandlerKnow
    Set ObjRs = ObjDml.ExeSelect(strSqlA, DbClass)
    If Not ObjRs.EOF And Not ObjRs.BOF Then
        Ret = ObjDml.ExeUpdate(strSql, DbClass)
        Set ObjDml = Nothing
        Set ObjRs = Nothing
    End If
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    If Not ObjRs Is Nothing Then
       Set ObjRs = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：修改FunctionRelation表信息
'返回值：无
'输入参数:FunctionID, fFunctionID
Public Function EditRelation(ByVal FunctionID As String, ByVal fFunctionID As String)
    Dim ID As String
    Dim ObjDml As Object
    Dim strSql, strSqlA, strSqlB, strSqlC As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim ObjRs As Recordset
    
    ErrorNo = 0
    '创建DML对象
    If FunctionID = "" And fFunctionID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    If fFunctionID = "" Then
        strSql = "delete from FunctionRelation where FunctionID='" & FunctionID & "'"
        On Error GoTo ErrorHandlerKnow
        Ret = ObjDml.ExeDelete(strSql, DbClass)
        Set ObjDml = Nothing
        GetObjectContext.SetComplete
        Exit Function
    End If
    strSqlA = "select * from FunctionRelation where FunctionID='" & FunctionID & "'"
    On Error GoTo ErrorHandlerKnow
    Set ObjRs = ObjDml.ExeSelect(strSqlA, DbClass)
    If ObjRs.EOF And ObjRs.BOF Then
        ID = AddRelation(FunctionID, fFunctionID)
    Else
        strSql = "Update FunctionRelation set fFunctionID='" & fFunctionID & "'" _
                & " where FunctionId='" & FunctionID & "'"
        '执行修改FunctionRelation表操作
        Ret = ObjDml.ExeUpdate(strSql, DbClass)
    End If
    ObjRs.Close
    Set ObjRs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    GetObjectContext.SetAbort
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjRs Is Nothing Then
        Set ObjRs = Nothing
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：修改FunctionLocale,FunctionRelation表信息
'返回值：无
'输入参数:FunctionID, fFunctionID,FunctionName,Locale
Public Function EditFunctionAll(ByVal FunctionID As String, ByVal fFunctionID As String, _
                                ByVal FunctionName As String, ByVal FunctionType As String, _
                                Optional ByVal Locale As String = "zh")
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    FunctionName = FiltQuoteChar(FunctionName)
    If FunctionID = "" And fFunctionID = "" And FunctionName = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    '执行修改FunctionLocale,FunctionRelation表操作
    Ret = EditFunction(FunctionID, FunctionType)
    Ret = EditFunctionLocale(FunctionID, FunctionName, Locale)
    Ret = EditRelation(FunctionID, fFunctionID)
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    GetObjectContext.SetAbort
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：修改FunctionLocale表信息
'返回值：无
'输入参数:FunctionID, Locale, FunctionName
Public Function EditFunctionLocale(ByVal FunctionID As String, ByVal FunctionName As String, _
                                    Optional ByVal Locale As String = "zh") As Long
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    FunctionName = FiltQuoteChar(FunctionName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "Update FunctionLocale set FunctionName='" & FunctionName & "'" _
            & " where FunctionId='" & FunctionID & "' and locale='" & Locale & "'"
    '执行修改FunctionLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：删除Function表信息
'返回值：无
'输入参数:FunctionID
Public Function DelFunction(ByVal FunctionID As String)
    Dim ObjDml As Object
    Dim strSqlA, strSqlB, strSqlC, strSqlD, strSql As String
    Dim RetA, RetB, RetC, RetD As Long
    Dim ErrNum As Long
    Dim Rs As Recordset
    Dim Description As String
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    
    On Error GoTo ErrorHandlerKnow
    '检查功能描述是否UserManage
    strSql = "select Description from [Function] where FunctionID='" & FunctionID & "'"
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    Description = Trim(Rs("Description"))
    If Description = AdminFunction Then
        ErrorNo = 10762
        GoTo ErrorHandlerKnow
    End If
    strSqlA = "Delete from UserFunction where FunctionID='" & FunctionID & "'"
    strSqlB = "Delete from FunctionLocale where FunctionID='" & FunctionID & "'"
    strSqlC = "Delete from GroupFunction where FunctionID='" & FunctionID & "'"
    strSqlD = "Delete from [Function] where FunctionID='" & FunctionID & "'"
    '执行删除Function表操作
    RetA = ObjDml.ExeDelete(strSqlA, DbClass)
    RetB = ObjDml.ExeDelete(strSqlB, DbClass)
    RetC = ObjDml.ExeDelete(strSqlC, DbClass)
    RetD = ObjDml.ExeDelete(strSqlD, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    GetObjectContext.SetAbort
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not Rs Is Nothing Then
        Set Rs = Nothing
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：删除FunctionLocale表信息
'返回值：无
'输入参数:FunctionID,Locale
Public Function DelFunctionLocale(ByVal FunctionID As String, _
                    Optional ByVal Locale As String = "zh")
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "delete from FunctionLocale where FunctionID ='" & FunctionID & "'" _
                & "and Locale='" & Locale & "'"
    '执行删除FunctionLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：删除FunctionRelation表信息
'返回值：无
'输入参数:FunctionID,fFunctionID
Public Function DelRelation(Optional ByVal FunctionID As String, Optional ByVal fFunctionID As String)

    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "delete from FunctionRelation where FunctionID ='" & FunctionID & "'" _
                & "and fFunctionID='" & fFunctionID & "'"
    '执行删除FunctionRelatione表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：获取添加Function时最大的OrderNum+1
'返回值：新纪录的OrderNum
'参数：无
Private Function GetMaxNum() As Integer
    Dim Rs As Recordset
    Dim strSql As String
    Dim ErrNum As Long
    Dim ObjDml As Object
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    strSql = "select * from [Function]"
    
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    If Rs.BOF And Rs.EOF Then
        GetMaxNum = 1
    Else
        strSql = "select Max(OrderNum) from [Function]"
        Set Rs = ObjDml.ExeSelect(strSql, DbClass)
        GetMaxNum = Rs(0) + 1
    End If
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not Rs Is Nothing Then
        Set Rs = Nothing
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：过滤字符串中的单引号
'返回值：过滤后的字符串
'参数：任意字符窜
Private Function FiltQuoteChar(strInput As String) As String
    Dim str1 As String
    Dim i As Integer
    Dim ErrNum As Long
    
    On Error GoTo ErrorHandlerKnow
    str1 = InStr(1, strInput, "'")
    If str1 <> 0 Then
        ErrNum = 10774
        GoTo ErrorHandlerKnow
    End If
    FiltQuoteChar = strInput
    Exit Function
    
ErrorHandlerKnow:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function

'功能：添加功能关系表数据
'返回值：ID
'参数：FunctionID,fFunctionID-父功能ID
Public Function AddRelation(ByVal FunctionID As String, ByVal fFunctionID As String) As String
    Dim ID As String
    Dim ObjDml As Object
    Dim ObjRs As Recordset
    Set ObjRs = CreateObject("Adodb.Recordset")
    Dim strSql, strSqlA As String
    Dim Ret As Long
    Dim ErrNum As Long
       
    ErrorNo = 0
    '创建DML对象
    If fFunctionID = "" Or FunctionID = "" Then
        Exit Function
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml1.ClsDml1")
    
    On Error GoTo ErrorHandlerKnow
    ID = ObjDml.GetGuid()
    strSql = "Insert into FunctionRelation (ID,FunctionID,fFunctionID) values " _
            & "('" & ID & "','" & FunctionID & "','" & fFunctionID & "')"
    '执行插入Function表操作
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    AddRelation = ID
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandlerKnow:
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnKnow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDml Is Nothing Then
        Set ObjDml = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description
End Function
