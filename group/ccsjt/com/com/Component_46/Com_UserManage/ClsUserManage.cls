VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ClsUserManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************
'* Author:LiangYe And Mr.wei     *
'* Modify Date:2000-11-01        *
'* Purpose:AMS用户管理            *
'*********************************
Option Explicit

Private Const DbClass = 7                      '数据库指向UserManage
Private ErrorNo As Long
Private Const AdminLogin = "Administrator"      '管理用户的Login账号
'Private sLocale As String                      '语言版本

''赋语言版本
'Public Property Let Locale(ByVal strLocale As String)
'    sLocale = strLocale
'End Property
'
''取语言版本
'Public Property Get Locale() As String
'    Locale = sLocale
'End Property
'
''赋属性初值
'Private Sub Class_Initialize()
'    sLocale = "zh"                         '语言版本
'End Sub

'*****************************************************************************

'功能:添加数据到用户功能表
'返回值:无
'输入参数：UserID,FunctionID
Public Function PutUserFunction(ByRef UserID, ByRef FunctionID)
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Ret As Long
    Dim ArrayNum As Integer
    Dim i As Integer
    Dim BCheck As Boolean           '检查记录是否已存在
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '输入参数不能都为数组
    If IsArray(UserID) And IsArray(FunctionID) Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    
    On Error GoTo ErrorHandlerKnow
    If IsArray(UserID) Then
        '用户ID为数组时，添加多个用户
        If FunctionID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(UserID)
        For i = 0 To ArrayNum
            '判断记录是否存在
            BCheck = CheckUserFunction(UserID(i), FunctionID)
            If BCheck = True Then
                strSql = "insert into UserFunction (UserID,FunctionID) values ('" & UserID(i) & "', " _
                            & "'" & FunctionID & "')"
                Ret = ObjDml.ExeInsert(strSql, DbClass)
            End If
        Next
    ElseIf IsArray(FunctionID) Then
        '功能ID为数组时，添加多个功能
        If UserID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(FunctionID)
        For i = 0 To ArrayNum
            BCheck = CheckUserFunction(UserID, FunctionID(i))
            If BCheck = True Then
                strSql = "insert into UserFunction (UserID,FunctionID) values ('" & UserID & "', " _
                        & "'" & FunctionID(i) & "')"
                Ret = ObjDml.ExeInsert(strSql, DbClass)
            End If
        Next
    Else
        If UserID = "" Or FunctionID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        BCheck = CheckUserFunction(UserID, FunctionID)
        If BCheck = True Then
            strSql = "insert into UserFunction (UserID,FunctionID) values ('" & UserID & "', " _
                        & "'" & FunctionID & "')"
            Ret = ObjDml.ExeInsert(strSql, DbClass)
        End If
    End If
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:添加数据到组功能表
'返回值:无
'输入参数：GroupID,FunctionID
Public Function PutGroupFunction(ByRef GroupID, ByRef FunctionID)
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Ret As Long
    Dim ArrayNum As Integer
    Dim i As Integer
    Dim BCheck As Boolean           '检查记录是否已存在
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '输入参数不能都为数组
    If IsArray(GroupID) And IsArray(FunctionID) Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    
    On Error GoTo ErrorHandlerKnow
    If IsArray(GroupID) Then
        '用户ID为数组时，添加多个用户
        If FunctionID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(GroupID)
        For i = 0 To ArrayNum
            BCheck = CheckGroupFunction(GroupID(i), FunctionID)
            If BCheck = True Then
                strSql = "Insert into GroupFunction (GroupID,FunctionID) values ('" & GroupID(i) & "', " _
                            & "'" & FunctionID & "')"
                Ret = ObjDml.ExeInsert(strSql, DbClass)
            End If
        Next
    ElseIf IsArray(FunctionID) Then
        '功能ID为数组时，添加多个功能
        If GroupID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(FunctionID)
        For i = 0 To ArrayNum
            BCheck = CheckGroupFunction(GroupID, FunctionID(i))
            If BCheck = True Then
                strSql = "insert into GroupFunction (GroupID,FunctionID) values ('" & GroupID & "', " _
                            & "'" & FunctionID(i) & "')"
                Ret = ObjDml.ExeInsert(strSql, DbClass)
            End If
        Next
    Else
        If GroupID = "" Or FunctionID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        BCheck = CheckGroupFunction(GroupID, FunctionID)
        If BCheck = True Then
            strSql = "insert into GroupFunction (GroupID,FunctionID) values ('" & GroupID & "', " _
                            & "'" & FunctionID & "')"
            Debug.Print strSql
            Ret = ObjDml.ExeInsert(strSql, DbClass)
        End If
    End If
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:添加数据到组功能表前检查此记录是否存在
'返回值:True或False
'输入参数：GroupID,FunctionID
Private Function CheckGroupFunction(ByVal GroupID As String, ByVal FunctionID As String) As Boolean
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "Select * from GroupFunction Where GroupId='" & GroupID & "' and" _
                & " FunctionId='" & FunctionID & "'"
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    If Rs.EOF And Rs.BOF Then
        CheckGroupFunction = True
    Else
        CheckGroupFunction = False
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:添加数据到用户功能表前检查此记录是否存在
'返回值:True或False
'输入参数：UserID,FunctionID
Private Function CheckUserFunction(ByVal UserID As String, ByVal FunctionID As String) As Boolean
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "Select * from UserFunction Where UserId='" & UserID & "' and" _
                & " FunctionId='" & FunctionID & "'"
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    If Rs.EOF And Rs.BOF Then
        CheckUserFunction = True
    Else
        CheckUserFunction = False
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:添加用户到组
'返回值:无
'输入参数：UserID,GroupID
Public Function PutUserGroup(ByRef UserID, ByRef GroupID)
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Ret As Long
    Dim ArrayNum As Integer
    Dim i As Integer
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '输入参数不能都为数组
    If IsArray(UserID) And IsArray(GroupID) Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    
    On Error GoTo ErrorHandlerKnow
    If IsArray(UserID) Then
        '用户ID为数组时，添加多个用户
        If GroupID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(UserID)
        For i = 0 To ArrayNum
            strSql = "Insert into UserGroup (UserID,GroupID) values ('" & UserID(i) & "', " _
                        & "'" & GroupID & "')"
            Ret = ObjDml.ExeInsert(strSql, DbClass)
        Next
    ElseIf IsArray(GroupID) Then
        '功能ID为数组时，添加多个功能
        If UserID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        ArrayNum = UBound(GroupID)
        For i = 0 To ArrayNum
            strSql = "insert into UserGroup (UserID,GroupID) values ('" & UserID & "', " _
                        & "'" & GroupID(i) & "')"
            Ret = ObjDml.ExeInsert(strSql, DbClass)
        Next
    Else
        If UserID = "" Or GroupID = "" Then
            ErrorNo = 10682
            GoTo ErrorHandlerKnow
        End If
        strSql = "insert into UserGroup (UserID,GroupID) values ('" & UserID & "', " _
                        & "'" & GroupID & "')"
        Ret = ObjDml.ExeInsert(strSql, DbClass)
    End If
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：删除GroupFunction表信息
'返回值：无
'输入参数: GroupID,FunctionID
Public Function DelGroupFunction(Optional ByVal GroupID As String, Optional ByVal FunctionID As String)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If GroupID = "" And FunctionID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    ElseIf GroupID = "" Then
        strSql = "Delete from GroupFunction where FunctionID='" & FunctionID & "'"
    ElseIf FunctionID = "" Then
        strSql = "Delete from GroupFunction where GroupID='" & GroupID & "'"
    ElseIf GroupID <> "" And FunctionID <> "" Then
        strSql = "Delete from GroupFunction where GroupID ='" & GroupID & "' And " _
                & "FunctionID='" & FunctionID & "'"
    End If
    '执行删除GroupFunction表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：删除UserFunction表信息
'返回值：无
'输入参数:UserID,FunctionID
Public Function DelUserFunction(Optional ByVal UserID As String, Optional ByVal FunctionID As String)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If UserID = "" And FunctionID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    ElseIf UserID = "" Then
        strSql = "Delete from UserFunction where FunctionID='" & FunctionID & "'"
    ElseIf FunctionID = "" Then
        strSql = "Delete from UserFunction where UserID='" & UserID & "'"
    ElseIf UserID <> "" And FunctionID <> "" Then
        strSql = "Delete from UserFunction where UserID ='" & UserID & "' And " _
                & "FunctionID='" & FunctionID & "'"
    End If
    '执行删除UserFunction表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:删除UserGroup记录
'返回值:无
'输入参数：UserID，GroupID
Public Function DelUserGroup(Optional ByVal UserID As String, Optional ByVal GroupID As String)
    Dim ObjDml As Object
    Dim strSql, strSqlFunc As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    
    On Error GoTo ErrorHandlerKnow
    If UserID = "" And GroupID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    ElseIf UserID = "" Then
        strSql = "Delete from UserGroup where GroupID='" & GroupID & "'"
        '执行删除UserFunction表操作
        strSqlFunc = "Delete from View_UserGroupFunction Where GroupID='" & GroupID & "'"
        Ret = ObjDml.ExeDelete(strSqlFunc, DbClass)
    ElseIf GroupID = "" Then
        strSql = "Delete from UserGroup where UserID='" & UserID & "'"
        '执行删除UserFunction表操作
        strSqlFunc = "select FunctionID from View_UserGroupFunction Where UserID='" & UserID & "'"
        Set Rs = ObjDml.ExeSelect(strSqlFunc, DbClass)
        Do While Not Rs.EOF
            Ret = DelUserFunction(UserID, Rs("FunctionID"))
            Rs.MoveNext
        Loop
    ElseIf UserID <> "" And GroupID <> "" Then
        strSql = "Delete from UserGroup where UserID ='" & UserID & "' And " _
                & "GroupID='" & GroupID & "'"
        '执行删除UserFunction表操作
        strSqlFunc = "select FunctionID from GroupFunction Where GroupID='" & GroupID & "'"
        Set Rs = ObjDml.ExeSelect(strSqlFunc, DbClass)
        Do While Not Rs.EOF
            Ret = DelUserFunction(UserID, Rs("FunctionID"))
            Rs.MoveNext
        Loop
    End If
    '执行删除UserGroup表操作
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:暂停用户
'返回值:无
'输入参数：UserID（UserInfo用户状态E:有效V:无效S:暂停---Update   UserHistory-Insert）
Public Function PauseUser(ByVal UserID As String)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim Status As String
    Dim UserHistory(3)
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    Status = "S"
    On Error GoTo ErrorHandlerKnow
    strSql = "Select LoginID,EndDate,Status from UserInfo where UserID='" & UserID & "'"
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    UserHistory(0) = UserID
    UserHistory(1) = Rs("LoginID")
    UserHistory(2) = Rs("EndDate")
    UserHistory(3) = Status
    '只有有效用户才可暂停
    If Rs("Status") = "E" And Rs("EndDate") > Now() Then
        strSql = "Update UserInfo set Status='" & Status & "' where UserID='" & UserID & "'"
        '执行修改UserInfo表操作
        Ret = ObjDml.ExeUpdate(strSql, DbClass)
        '添加历史纪录
        Ret = AddUserHistory(UserHistory)
    Else
        ErrorNo = 10709
        GoTo ErrorHandlerKnow
    End If
    
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:将暂停的用户恢复
'返回值:无
'输入参数：UserID（UserInfo用户状态E:有效V:无效S:暂停---Update   UserHistory-Insert）
Public Function RestoreUser(ByVal UserID As String)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim Status As String
    Dim UserHistory(3)
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    Status = "E"
    strSql = "Select LoginID,EndDate,Status from UserInfo where UserID='" & UserID & "'"
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    UserHistory(0) = UserID
    UserHistory(1) = Rs("LoginID")
    UserHistory(2) = Rs("EndDate")
    UserHistory(3) = Status
    '只有暂停用户才可恢复
    If Rs("Status") = "S" And Rs("EndDate") > Now() Then
        strSql = "Update UserInfo set Status='" & Status & "' where UserID='" & UserID & "'"
        '执行修改UserInfo表操作
        Ret = ObjDml.ExeUpdate(strSql, DbClass)
        '添加历史纪录
        Ret = AddUserHistory(UserHistory)
    Else
        ErrorNo = 10711
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:判断LoginID是否存在
'返回值:布尔变量(如果存在此用户返回True,否则返回False)
'输入参数：LoginID
Public Function IsExistLoginID(ByVal LoginID As String) As Boolean
    Dim ObjDml As Object
    Dim strSql As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If LoginID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    strSql = "Select LoginID from LoginInfo where LoginID='" & LoginID & "'"
    '执行UserInfo表Select操作
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If Rs.EOF And Rs.BOF Then
        IsExistLoginID = False
    Else
        IsExistLoginID = True
        ErrorNo = 10706
        GoTo ErrorHandlerKnow
    End If
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:判断LoginID和Password是否正确
'返回值:布尔变量(如果校验用户正确返回True,否则返回False)
'输入参数：LoginID,Password,UseObject
'Public Function CheckLoginID(ByVal LoginID As String, ByVal Password As String, _
'                           ByVal UseObject As String) As Boolean
Public Function CheckLoginID(ByVal LoginID As String, ByVal Password As String, _
                           ByVal UseObject As String) As Boolean
    Dim ObjDml As Object
    Dim strSql As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    LoginID = FiltQuoteChar(LoginID)
    Password = FiltQuoteChar(Password)
    
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If LoginID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    strSql = "Select Password,Status,EndDate from View_UserInfo where LoginID='" & LoginID & "'" _
                & " and UseObject='" & UseObject & "'"
    '执行UserInfo表Select操作
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    'Set Rs = ObjDml.ExeSQL(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    '校验登录用户是否有效
    If Rs.EOF And Rs.BOF Then
        CheckLoginID = False
        ErrorNo = 999999 '10707
        GoTo ErrorHandlerKnow
    Else
        If Password = Rs(0) Then
            '只有有效用户在结束时间之前才可登录
            If Rs("Status") = "E" And Rs("EndDate") > Now() Then
                CheckLoginID = True
            Else
                CheckLoginID = False
                ErrorNo = 10767
               GoTo ErrorHandlerKnow
            End If
        Else
            CheckLoginID = False
            ErrorNo = 10707
            GoTo ErrorHandlerKnow
        End If
    End If
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:判断用户是否有此功能
'返回值:布尔变量(如果用户功能存在返回True,否则返回False)
'输入参数：UserID,FunctionDescription
Public Function VerifyUserFunction(ByVal UserID As String, ByVal FunctionDescription As String) As Boolean
    Dim ObjDml As Object
    Dim strSql As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    Dim RsGroup As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If UserID = "" Or FunctionDescription = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    strSql = "Select UserID from View_UserFunction where UserID='" & UserID & "'" _
                & " And Description='" & FunctionDescription & "'"
    '执行UserInfo表Select操作
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If Rs.EOF And Rs.BOF Then
        VerifyUserFunction = False
    Else
        VerifyUserFunction = True
        Exit Function
    End If
    
    '取用户所属组，并校验该组有无此功能
    Set RsGroup = GetUserGroup(UserID)
    Do While Not RsGroup.EOF
        VerifyUserFunction = VerifyGroupFunction(RsGroup("GroupID"), FunctionDescription)
        If VerifyUserFunction = True Then
            Exit Function
        End If
        RsGroup.MoveNext
    Loop
    Set RsGroup = Nothing
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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
    If Not RsGroup Is Nothing Then
        Set RsGroup = Nothing
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

'功能:判断组是否有此功能
'返回值:布尔变量(如果组功能存在返回True,否则返回False)
'输入参数：GroupID,FunctionDescription
Public Function VerifyGroupFunction(ByVal GroupID As String, ByVal FunctionDescription As String) As Boolean
    Dim ObjDml As Object
    Dim strSql As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    '创建DML对象
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    If GroupID = "" Or FunctionDescription = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    strSql = "Select GroupID from View_GroupFunction where GroupID='" & GroupID & "'" _
                & " And Description='" & FunctionDescription & "'"
    '执行UserInfo表Select操作
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If Rs.EOF And Rs.BOF Then
        VerifyGroupFunction = False
    Else
        VerifyGroupFunction = True
    End If
    Set Rs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'*****************************************************************************

'功能:取代理商记录集
'返回值:AgentInfo记录集
'输入参数：无
Public Function GetAgent() As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    '取AgentInfo中DealStatus="V"(有效)的记录
    strSql = "select * from AgentInfo where DealStatus='V'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetAgent = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetAgent.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:取Company记录集
'返回值:View_Company记录集
'输入参数：Locale
Public Function GetCompany(ByVal Locale As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from View_Company where Locale='" & Locale & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetCompany = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetCompany.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:取语言版本记录集
'返回值:Locale记录集
'输入参数：无
Public Function GetLocale() As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from Locale"
    
    On Error GoTo ErrorHandlerKnow
    Set GetLocale = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetLocale.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:取用户对象记录集
'返回值:UseObject记录集
'输入参数：无
Public Function GetUseObject() As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long

    ErrorNo = 0

    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from UseObject"

    On Error GoTo ErrorHandlerKnow
    Set GetUseObject = ObjDml.ExeSelect(strSql, DbClass)

    On Error GoTo ErrorHandlerUnKnow
    If GetUseObject.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:获取无效的登录用户View_VolidLogin记录集
'返回值:View_VolidLogin记录集
'输入参数：UseObject
Public Function GetVolidLogin(ByVal UseObject As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from View_VolidLogin where UseObject='" & UseObject & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetVolidLogin = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据LoginID获取LoginInfo记录集
'返回值:LoginInfo记录集
'输入参数：LoginID,UseObject
Public Function GetLoginInfo(ByVal LoginID As String, ByVal UseObject As String) As Recordset
    
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If LoginID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from View_UserInfo where LoginID='" & LoginID & "' and UseObject='" & UseObject & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetLoginInfo = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetLoginInfo.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取UserHistory记录集
'返回值:UserHistory记录集
'输入参数：UserID
Public Function GetUserHistory(ByVal UserID As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from UserHistory where UserID='" & UserID & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetUserHistory = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetUserHistory.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定视图View_UserInfo记录集
'返回值:View_UserInfo记录集
'输入参数：UserID,Locale,UseObject
Public Function GetUserInfo(ByVal UserID As String, ByVal Locale As String, _
                            ByVal UseObject As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_UserInfo where UserID='" & UserID & "' and " _
                    & "UseObject='" & UseObject & "'"
    Else
        strSql = "select * from View_UserInfo where UserID='" & UserID & "'" _
                & " and locale='" & Locale & "' and UseObject='" & UseObject & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetUserInfo = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetUserInfo.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定视图View_ComputerInfo(机器)记录集
'返回值:View_ComputerInfo记录集(View_ComputerInfo与View_UserInfo只是UserType不同)
'输入参数：UserID,Locale,UseObject
Public Function GetComputerInfo(ByVal UserID As String, ByVal Locale As String, _
                            ByVal UseObject As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_ComputerInfo where UserID='" & UserID & "' and " _
                    & "UseObject='" & UseObject & "'"
    Else
        strSql = "select * from View_ComputerInfo where UserID='" & UserID & "'" _
                & " and locale='" & Locale & "' and UseObject='" & UseObject & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetComputerInfo = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetComputerInfo.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取AgentID
'返回值:AgentID
'输入参数：UserID
Public Function UserIDToAgentID(ByVal UserID As String) As String
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    If UserID = "" Then
        UserIDToAgentID = ""
        Exit Function
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    'strSql = "select Distinct AgentID From View_UserInfo Where Convert(UserID)='" & UserID & "'"
    strSql = "select Distinct AgentID From View_UserInfo Where '{'+Convert(Char(36),UserID)+'}'='" & UserID & "'"
    On Error GoTo ErrorHandler
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If Rs.EOF Then
        GoTo ErrorHandler
    ElseIf IsNull(Rs("AgentID")) Then
        GoTo ErrorHandler
    End If
    UserIDToAgentID = Rs("AgentID")
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandler:
    Set Rs = Nothing
    Set ObjDml = Nothing
    UserIDToAgentID = ""
    Exit Function
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定View_GroupInfo视图信息记录集
'返回值:用户信息记录集
'输入参数：GroupID,Locale
Public Function GetGroupInfo(ByVal GroupID As String, Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    ErrorNo = 0
    If GroupID = "" Then
        ErrorNo = 10682
         GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_GroupInfo where GroupID='" & GroupID & "'"
    Else
        strSql = "select * from View_GroupInfo where GroupID='" & GroupID & "'" _
                & " and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set GetGroupInfo = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If GetGroupInfo.EOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据GroupID获取指定GroupLocale信息记录集
'返回值:组Locale信息记录集
'输入参数：GroupID
Public Function GetGroupLocale(ByVal GroupID As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    ErrorNo = 0
    If GroupID = "" Then
        ErrorNo = 10682
         GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from GroupLocale where GroupID='" & GroupID & "'"
    On Error GoTo ErrorHandlerKnow
    Set GetGroupLocale = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据GroupID,Locale获取指定组名GroupName
'返回值:指定语言组名
'输入参数：GroupID,Locale
Public Function GetGroupName(ByVal GroupID As String, ByVal Locale As String) As String
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    If GroupID = "" Or Locale = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from GroupLocale where GroupID='" & GroupID & "'" _
                & " and Locale='" & Locale & "'"
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    GetGroupName = Rs("GroupName")
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:从View_UserInfo取指定语言所有记录的记录集
'返回值:View_UserInfo记录集
'输入参数：Locale,UseObject
Public Function GetAllUser(ByVal Locale As String, ByVal UseObject As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If Locale = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    
    strSql = "select * from View_UserInfo where locale='" & Locale & "' and UseObject='" & UseObject & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetAllUser = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:从View_ComputerInfo取指定语言所有记录的记录集
'返回值:View_ComputerInfo记录集
'输入参数：Locale,UseObject
Public Function GetAllComputer(ByVal Locale As String, ByVal UseObject As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If Locale = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    
    strSql = "select * from View_ComputerInfo where locale='" & Locale & "' and UseObject='" & UseObject & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetAllComputer = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据Locale获取指定View_GroupInfo视图信息记录集
'返回值:用户信息记录集
'输入参数：Locale
Public Function GetAllGroup(ByVal Locale As String) As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If Locale = "" Then
        ErrorNo = 10682
         GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select * from View_GroupInfo where Locale='" & Locale & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetAllGroup = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
'    If GetAllGroup.EOF Then
'        ErrorNo = 10684
'        GoTo ErrorHandlerKnow
'    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定视图View_UserGroup信息记录集
'返回值:信息记录集
'输入参数：UserID,locale
Public Function GetUserGroup(ByVal UserID, Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_UserGroup where UserID='" & UserID & "'"
    Else
        strSql = "select * from View_UserGroup where UserID='" & UserID & "'" _
                & " and locale='" & Locale & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetUserGroup = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：修改LoginInfo表信息
'返回值：无
'输入参数:LoginID
'       LoginInfo(4)
'       (0)Name (1)Sex(M或F) (2)AgentID(3)CompanyID (4)ContactInfo
Public Function EditLogin(ByVal LoginID As String, ByRef LoginInfo)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    '校验Name和Contactinfo中的"'"
    LoginInfo(0) = FiltQuoteChar(CStr(LoginInfo(0)))
    LoginInfo(4) = FiltQuoteChar(CStr(LoginInfo(4)))
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Update LoginInfo set Name='" & LoginInfo(0) & "', " _
            & "Sex='" & LoginInfo(1) & "',AgentID='" & LoginInfo(2) & "'," _
            & "CompanyID='" & LoginInfo(3) & "',ContactInfo='" & LoginInfo(4) & "'" _
            & " where LoginId='" & LoginID & "'"
    '执行修改LoginInfo表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：修改GroupLocale表信息
'返回值：无
'输入参数:GroupID
'       Locale
'       GroupName
Public Function EditGroupLocale(ByVal GroupID As String, ByVal GroupName As String, _
                                Optional ByVal Locale As String = "zh")
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    GroupName = FiltQuoteChar(GroupName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Update GroupLocale set GroupName='" & GroupName & "'" _
            & " where GroupId='" & GroupID & "' and locale='" & Locale & "'"
    '执行修改GroupLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：修改UserInfo表EndDate
'返回值：无
'输入参数:UserID
'       Password
Public Function EditUserEndDate(ByVal UserID As String, ByVal EndDate As Date)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    If Not IsDate(EndDate) Then
        ErrorNo = 10768
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Update UserInfo set EndDate='" & EndDate & "' where UserId='" & UserID & "'"
    '执行修改UserHistory表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Ret = EditUserHistory(UserID, EndDate)
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：修改UserInfo表信息
'返回值：无
'输入参数:UserID
'       Password
Public Function EditPassword(ByVal UserID As String, ByVal Password As String)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    Password = FiltQuoteChar(Password)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Update UserInfo set Password='" & Password & "'" _
            & " where UserId='" & UserID & "'"
    '执行修改UserInfo表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：删除GroupLocale表信息
'返回值：无
'输入参数:GroupID,Locale
Public Function DelGroupLocale(ByVal GroupID As String, Optional ByVal Locale As String = "zh")
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "delete from GroupLocale where GroupID ='" & GroupID & "'" _
                & " and Locale='" & Locale & "'"
    '执行删除GroupLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：删除GroupLocale,GroupFunction,UserGroup,GroupInfo表信息
'返回值：无
'输入参数:GroupID
Public Function DelGroup(ByVal GroupID As String)
    Dim ObjDml As Object
    Dim strSqlA, strSqlB, strSqlC, strSqlD As String
    Dim RetA, RetB, RetC, RetD As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSqlA = "Delete from GroupLocale where GroupID='" & GroupID & "'"
    strSqlB = "Delete from UserGroup where GroupID='" & GroupID & "'"
    strSqlC = "Delete from  GroupFunction where GroupID='" & GroupID & "'"
    strSqlD = "Delete from GroupInfo where GroupID='" & GroupID & "'"
    '执行删除GroupLocale,GroupFunction,UserGroup,GroupInfo表操作
    On Error GoTo ErrorHandlerKnow
    RetA = ObjDml.ExeDelete(strSqlA, DbClass)
    RetB = ObjDml.ExeDelete(strSqlB, DbClass)
    RetC = ObjDml.ExeDelete(strSqlC, DbClass)
    RetD = ObjDml.ExeDelete(strSqlD, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：删除UserFunction,UserGroup,UserInfo表信息
'返回值：无
'输入参数:UserID
Public Function DelUser(ByVal UserID As String)
    Dim ObjDml As Object
    Dim strSqlA, strSqlB, strSqlC, strSql As String
    Dim LoginID As String
    Dim RetA, RetB, RetC, Ret As Long
    Dim ErrNum As Long
    Dim Rs As Recordset
    Dim UserHistory(3)
    Dim EndDate As Date
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    
    On Error GoTo ErrorHandlerKnow
    '检查是否Administrator
    strSql = "select LoginID,EndDate from UserInfo where UserID='" & UserID & "'"
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    LoginID = Rs("LoginID")
    EndDate = Rs("EndDate")
    If LoginID = AdminLogin Then
        ErrorNo = 10761
        GoTo ErrorHandlerKnow
    End If
    
    strSqlA = "Delete from UserFunction where UserID='" & UserID & "'"
    strSqlB = "Delete from UserGroup where UserID='" & UserID & "'"
    strSqlC = "Delete from  UserInfo where UserID='" & UserID & "'"
    '执行删除UserFunction,UserGroup,UserInfo表操作
    RetA = ObjDml.ExeDelete(strSqlA, DbClass)
    RetB = ObjDml.ExeDelete(strSqlB, DbClass)
    RetC = ObjDml.ExeDelete(strSqlC, DbClass)
    
    '添加历史纪录
    UserHistory(0) = UserID
    UserHistory(1) = LoginID
    UserHistory(2) = EndDate
    UserHistory(3) = "V"
    Ret = AddUserHistory(UserHistory)
    
    On Error GoTo ErrorHandlerUnKnow
    Set Rs = Nothing
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：修改UserHistory信息
'返回值：无
'输入参数:UserID,EndDate
Private Function EditUserHistory(ByVal UserID As String, ByVal EndDate As Date)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim Rs As Recordset
    Dim StartDate As Date
    Dim StartYear As String
    Dim StartMonth As String
    Dim StartDay As String
    Dim StartTime As String
    Dim StartTime1 As String
        
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    If Not IsDate(EndDate) Then
        ErrorNo = 10768
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "select Max(StartDate) from UserHistory where UserId='" & UserID & "'"
    
    On Error GoTo ErrorHandlerKnow
    '查询历史纪录中StartDate最晚的记录
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    If IsNull(Rs(0)) Then
        Exit Function
    End If
    StartDate = Rs(0)
    StartYear = Year(StartDate)
    StartMonth = Month(StartDate)
    StartDay = Day(StartDate)
    StartTime = Format(StartDate, "hh:mm:ss")
    strSql = "Update UserHistory set EndDate='" & EndDate & "' where UserId='" & UserID & "'" _
            & " And year(StartDate)='" & StartYear & "' and month(StartDate)='" & StartMonth & "'" _
            & "and day(StartDate)='" & StartDay & "' and convert(Char,StartDate,8)='" & StartTime & "'"
    '执行修改UserHistory表操作
    Ret = ObjDml.ExeUpdate(strSql, DbClass)
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据GroupID,UseObject和Locale获取指定视图信息记录集
'返回值:信息记录集
'输入参数：GroupID,UseObject,Locale
Public Function GetGroupUser(ByVal GroupID As String, ByVal UseObject As String, _
                            Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If GroupID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_GroupUser where GroupID='" & GroupID & "' and " _
                    & "UseObject='" & UseObject & "'"
    Else
        strSql = "select * from View_GroupUser where GroupID='" & GroupID & "'" _
                & " and locale='" & Locale & "' and UseObject='" & UseObject & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetGroupUser = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定视图View_UserFunction信息记录集
'返回值:信息记录集
'输入参数：UserID,Locale
Public Function GetUserFunction(ByVal UserID As String, _
                Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_UserFunction where UserID='" & UserID & "'"
    Else
        strSql = "select * from View_UserFunction where UserID='" & UserID & "'" _
                & " and locale='" & Locale & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetUserFunction = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据UserID获取指定视图View_FuncStr信息记录集的OrderNum字符串
'返回值:OrderNum字符串
'输入参数：UserID
Public Function GetFuncStr(ByVal UserID As String) As String
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim CompResult As Long          '字符串比较结果
    Dim Rs As Recordset
    Dim RsGroup As Recordset
    Dim StrFunc As String
    Dim str1, Str2 As String
    
    ErrorNo = 0
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    strSql = "select OrderNum from View_FuncStr where UserID='" & UserID & "'" _
                & " Order by OrderNum Asc"
    
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    
    StrFunc = ","
    Do While Not Rs.EOF
        StrFunc = StrFunc & CStr(Rs("OrderNum")) & ","
        Rs.MoveNext
    Loop
    '取用户所属组，并加上该组拥有的功能OrderNum，相同的功能要过滤
    On Error GoTo ErrorHandlerKnow
    strSql = "Select * from View_GroupFuncStr where UserID='" & UserID & "'"
    Set RsGroup = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Do While Not RsGroup.EOF
        str1 = CStr(RsGroup("OrderNum"))
        Str2 = "," & str1 & ","
        CompResult = InStr(1, StrFunc, Str2)
        If CompResult = 0 Then
            StrFunc = StrFunc & str1 & ","
        End If
        RsGroup.MoveNext
    Loop
    GetFuncStr = StrFunc
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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
    If Not RsGroup Is Nothing Then
        Set RsGroup = Nothing
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

'功能:根据UserID和ComputerID获取共同的功能OrderNum字符串
'返回值:OrderNum字符串
'输入参数：UserID,ComputerID
Public Function GetFuncTogether(ByVal UserID As String, ByVal ComputerID As String) As String
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    Dim CompResult As Long          '字符串比较结果
    Dim Rs As Recordset
    Dim RsGroup As Recordset
    Dim StrFunc, StrUser As String
    Dim str1, Str2 As String
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    '取用户的功能
    StrUser = GetFuncStr(UserID)
    If UserID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    '取计算机自身具有的功能，与用户功能相同的保留
    strSql = "Select OrderNum from View_FuncStr where UserID='" & ComputerID & "'" _
                & " Order by OrderNum Asc"
    
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    StrFunc = ","
    Do While Not Rs.EOF
        str1 = CStr(Rs("OrderNum"))
        Str2 = "," & str1 & ","
        CompResult = InStr(1, StrUser, Str2)
        If CompResult <> 0 Then
            StrFunc = StrFunc & str1 & ","
        End If
        Rs.MoveNext
    Loop
    '取计算机所属组，并加上该组拥有的功能OrderNum
    On Error GoTo ErrorHandlerKnow
    strSql = "Select * from View_GroupFuncStr where UserID='" & ComputerID & "'"
    Set RsGroup = ObjDml.ExeSelect(strSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    Do While Not RsGroup.EOF
        str1 = CStr(RsGroup("OrderNum"))
        Str2 = "," & str1 & ","
        CompResult = InStr(1, StrUser, Str2)
        If CompResult <> 0 Then
            StrFunc = StrFunc & str1 & ","
        End If
        RsGroup.MoveNext
    Loop
    GetFuncTogether = StrFunc
    Set Rs = Nothing
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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
    If Not RsGroup Is Nothing Then
        Set RsGroup = Nothing
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

'功能:根据GroupID和指定语言获取指定视图信息记录集
'返回值:信息记录集
'输入参数：GroupID
Public Function GetGroupFunction(ByVal GroupID As String, Optional ByVal Locale As String = "zh") As Recordset
    Dim strSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If GroupID = "" Then
        ErrorNo = 10682
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        strSql = "select * from View_GroupFunction where GroupID='" & GroupID & "'"
    Else
        strSql = "select * from View_GroupFunction where GroupID='" & GroupID & "'" _
                    & " and locale='" & Locale & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetGroupFunction = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加GroupLocale表
'返回值：无
'输入参数:
'       GroupID
'       GroupName
'       Locale
Public Function AddGroupLocale(ByVal GroupID As String, ByVal GroupName As String, _
                                Optional ByVal Locale As String = "zh")
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    GroupName = FiltQuoteChar(GroupName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Insert into GroupLocale (GroupID,GroupName,Locale) values " _
            & "('" & GroupID & "','" & GroupName & "','" & Locale & "')"
    '执行插入GroupLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加新纪录到GroupInfo表
'返回值：GroupID
'输入参数:Description
Public Function AddGroup(ByVal Description As String) As String
    Dim GroupID As String
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    Description = FiltQuoteChar(Description)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    GroupID = ObjDml.GetGuid()
    strSql = "Insert into GroupInfo (GroupID,Description) values " _
            & "('" & GroupID & "','" & Description & "')"
    '执行插入Groupinfo表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    AddGroup = GroupID
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加到UserInfo,LoginInfo,UserHistory表
'返回值：UserID
'输入参数:
'       UserInfo(8)
'       LoginID    (0)
'       Name       (1)
'       Sex        (2)
'       AgentID    (3)
'       CompanyID  (4)
'       ContactInfo(5)
'       UseObject  (6)
'       password   (7)
'       EndDate    (8)
Public Function AddUser(ByRef UserInfo) As String
    Dim UserID As String
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim LoginInfo(6)
    Dim UserHistory(3)
    Dim i As Integer
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    UserInfo(0) = FiltQuoteChar(CStr(UserInfo(0)))
    UserInfo(1) = FiltQuoteChar(CStr(UserInfo(1)))
    UserInfo(5) = FiltQuoteChar(CStr(UserInfo(5)))
    UserInfo(7) = FiltQuoteChar(CStr(UserInfo(7)))
    On Error GoTo ErrorHandlerUnKnow
    If Not IsDate(UserInfo(8)) Then
        ErrorNo = 10768
        GoTo ErrorHandlerKnow
    End If
    '创建DML对象
    UserInfo(8) = CDate(UserInfo(8))
    If UserInfo(8) < Now() Then
        ErrorNo = 10710
        GoTo ErrorHandlerKnow
    End If
    For i = 0 To 6
        LoginInfo(i) = UserInfo(i)
    Next
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    UserID = ObjDml.GetGuid()
    UserHistory(0) = UserID
    UserHistory(1) = UserInfo(0)
    UserHistory(2) = UserInfo(8)
    UserHistory(3) = "E"
    strSql = "Insert into UserInfo (UserID,LoginID,password,EndDate) values " _
            & "('" & UserID & "','" & UserInfo(0) & "','" & UserInfo(7) & "','" & UserInfo(8) & "')"
    '执行插入UserInfo,LoginInfo,UserHistory表操作
    On Error GoTo ErrorHandlerKnow
    Ret = Addlogin(LoginInfo)
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Ret = AddUserHistory(UserHistory)
    AddUser = UserID
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加到UserInfo,LoginInfo,UserHistory表
'返回值：UserID
'输入参数:
'       ComputerInfo(8)
'       LoginID    (0)
'       Name       (1)
'       Sex        (2)
'       AgentID    (3)
'       CompanyID  (4)
'       ContactInfo(5)
'       UseObject  (6)
'       password   (7)
'       EndDate    (8)
Public Function AddComputer(ByRef ComputerInfo) As String
    Dim UserID As String
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim LoginInfo(6)
    Dim UserHistory(3)
    Dim i As Integer
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    ComputerInfo(0) = FiltQuoteChar(CStr(ComputerInfo(0)))
    ComputerInfo(1) = FiltQuoteChar(CStr(ComputerInfo(1)))
    ComputerInfo(5) = FiltQuoteChar(CStr(ComputerInfo(5)))
    ComputerInfo(7) = FiltQuoteChar(CStr(ComputerInfo(7)))
    On Error GoTo ErrorHandlerUnKnow
    If Not IsDate(ComputerInfo(8)) Then
        ErrorNo = 10768
        GoTo ErrorHandlerKnow
    End If
    '创建DML对象
    ComputerInfo(8) = CDate(ComputerInfo(8))
    If ComputerInfo(8) < Now() Then
        ErrorNo = 10710
        GoTo ErrorHandlerKnow
    End If
    For i = 0 To 6
        LoginInfo(i) = ComputerInfo(i)
    Next
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    UserID = ObjDml.GetGuid()
    UserHistory(0) = UserID
    UserHistory(1) = ComputerInfo(0)
    UserHistory(2) = ComputerInfo(8)
    UserHistory(3) = "E"
    strSql = "Insert into UserInfo (UserID,LoginID,password,EndDate,UserType) values " _
            & "('" & UserID & "','" & ComputerInfo(0) & "','" & ComputerInfo(7) & "','" & ComputerInfo(8) & "','C')"
    '执行插入UserInfo,LoginInfo,UserHistory表操作
    On Error GoTo ErrorHandlerKnow
    Ret = Addlogin(LoginInfo)
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Ret = AddUserHistory(UserHistory)
    AddComputer = UserID
    Set ObjDml = Nothing
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加新Login用户
'返回值：无
'输入参数:
'       Logininfo(6)
'       LoginID
'       Name
'       Sex
'       AgentID
'       CompanyID
'       Contactinfo
'       UseObject
Private Function Addlogin(ByRef LoginInfo)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim IsExist As Boolean
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    LoginInfo(0) = FiltQuoteChar(CStr(LoginInfo(0)))
    LoginInfo(1) = FiltQuoteChar(CStr(LoginInfo(1)))
    LoginInfo(5) = FiltQuoteChar(CStr(LoginInfo(5)))
    '检查用户是否存在，如果存在此LoginID则出错
    IsExist = IsExistLoginID(LoginInfo(0))
    
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSql = "Insert into logininfo (loginID,Name,Sex,AgentId," _
            & "CompanyID,ContactInfo,UseObject) values " _
            & "('" & LoginInfo(0) & "','" & LoginInfo(1) & "'," _
            & "'" & LoginInfo(2) & "','" & LoginInfo(3) & "','" _
            & LoginInfo(4) & "','" & LoginInfo(5) & "','" & LoginInfo(6) & "')"
    '执行插入LoginInfo表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加用户历史纪录
'返回值：无
'输入参数:
'       UserHistory(2)
'       0-UserID
'       1-LoginID
'       2-EndDate
'       3-Status
Private Function AddUserHistory(ByRef UserHistory)
    Dim ObjDml As Object
    Dim strSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim StartDate As Date
    
    ErrorNo = 0
    UserHistory(2) = CDate(UserHistory(2))
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    UserHistory(1) = FiltQuoteChar(CStr(UserHistory(1)))
    
    On Error GoTo ErrorHandlerUnKnow
    If Not IsDate(UserHistory(2)) Then
        ErrorNo = 10768
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StartDate = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    On Error GoTo ErrorHandlerKnow
    '修改上次最后一条记录的EndDate
    Ret = EditUserHistory(UserHistory(0), StartDate)
    '执行插入UserHistory表操作
    strSql = "Insert into UserHistory (UserID,LoginID,StartDate,EndDate,Status) values " _
            & "('" & UserHistory(0) & "','" & UserHistory(1) & "','" & StartDate & "'," _
            & "'" & UserHistory(2) & "','" & UserHistory(3) & "')"
    Ret = ObjDml.ExeInsert(strSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'************************************************************
'功能:根据条件查询视图View_UserInfo信息记录集
'返回值:用户信息记录集
'输入参数：UseObject用户对象  TextFileld:字段  TiaoJian:条件值
Public Function SearchUserInfo(ByVal UseObject As String, Optional ByVal TextField As String = "", _
                        Optional ByVal TiaoJian As String = "", Optional ByVal Locale As String = "zh") As Recordset
    Dim ObjDml As Object
    Dim strSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If TextField = "" Or TiaoJian = "" Then
        strSql = "select * from View_UserInfo where useobject='" & UseObject & "'" & " and Locale='" & Locale & "'"
'        strSql = "select * from View_UserInfo where useobject='" & UseObject & "'" _
'                    & " and Locale='" & Locale & "'"
    Else
        TextField = "%" & TextField & "%"
        strSql = "select * from View_UserInfo where useobject='" & UseObject & "'" & " and " _
                    & TiaoJian & " like '" & TextField & "' and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set SearchUserInfo = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchUserInfo.EOF And SearchUserInfo.BOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据条件查询视图View_ComputerInfo信息记录集
'返回值:用户信息记录集
'输入参数：UseObject用户对象  TextFileld:字段  TiaoJian:条件值
Public Function SearchComputerInfo(ByVal UseObject As String, ByVal TextField As String, _
                        ByVal TiaoJian As String, Optional ByVal Locale As String = "zh") As Recordset
    Dim ObjDml As Object
    Dim strSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If TextField = "" Or TiaoJian = "" Then
        strSql = "select * from View_ComputerInfo where useobject='" & UseObject & "'" _
                    & " and Locale='" & Locale & "'"
    Else
        TextField = "%" & TextField & "%"
        strSql = "select * from View_ComputerInfo where useobject='" & UseObject & "'" & " and " _
                    & TiaoJian & " like '" & TextField & "' and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set SearchComputerInfo = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchComputerInfo.EOF And SearchComputerInfo.BOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能:根据条件查询视图View_GroupInfo信息记录集
'返回值:组信息记录集
'输入参数：TextFileld:字段  TiaoJian:条件值
Public Function SearchGroupInfo(ByVal TextField As String, ByVal TiaoJian As String, _
                                Optional ByVal Locale As String = "zh") As Recordset
    Dim ObjDml As Object
    Dim strSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If TextField = "" Or TiaoJian = "" Then
        strSql = "select * from View_groupInfo where Locale='" & Locale & "'"
    Else
        TextField = "%" & TextField & "%"
        strSql = "select * from View_groupInfo where " & TiaoJian & " like '" & TextField & "'" _
                & " and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set SearchGroupInfo = ObjDml.ExeSelect(strSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchGroupInfo.EOF And SearchGroupInfo.BOF Then
        ErrorNo = 10684
        GoTo ErrorHandlerKnow
    End If
    Set ObjDml = Nothing
    Exit Function

ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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

'功能：添加新纪录到GroupInfo,GroupLocale表
'返回值：GroupID
'输入参数 Description
'         GroupName
'         Locale
Public Function AddGroupAll(ByVal Description As String, ByVal GroupName As String, _
                                Optional ByVal Locale As String = "zh") As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim GroupID As String
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    GroupName = FiltQuoteChar(GroupName)
    Description = FiltQuoteChar(Description)
    GroupID = AddGroup(Description)
    Ret = AddGroupLocale(GroupID, GroupName, Locale)
    AddGroupAll = GroupID
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
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
