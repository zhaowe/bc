VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ClsSzdm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************
'* Author:Mr.wei              *
'* Modify Date:2000-11-8    *
'* Purpose:三子码管理          *
'******************************
Option Explicit
Private Const DbClass = 9     '数据库指向Parameter
Private ErrorNo As Long

'功能:根据条件查询视图View_CityInfo信息记录集
'返回值:组信息记录集
'输入参数：TextFileld:字段  TiaoJian:条件值
Public Function SearchCityInfo(ByVal TextField As String, ByVal TiaoJian As String, ByVal Locale As String) As Recordset
    Dim ObjDml As Object
    Dim StrSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    On Error GoTo ErrorHandlerKnow
    If TextField = "" Or TiaoJian = "" Then
        StrSql = "select * from View_CityInfo where locale='" & Locale & "' order by cityid"
    Else
        TextField = FiltQuoteChar(TextField)
        TextField = "%" & TextField & "%"
        StrSql = "select * from View_CityInfo where " & TiaoJian & " like '" & TextField & "'" _
        & " and locale='" & Locale & "' order by cityid "
    End If
    Set SearchCityInfo = ObjDml.ExeSelect(StrSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchCityInfo.EOF And SearchCityInfo.BOF Then
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

'功能:根据条件查询视图View_AirportSzdmInfo信息记录集
'返回值:组信息记录集
'输入参数：TextFileld:字段  TiaoJian:条件值
Public Function SearchAirportInfo(ByVal TextField As String, ByVal TiaoJian As String, ByVal Locale As String) As Recordset
    Dim ObjDml As Object
    Dim StrSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    On Error GoTo ErrorHandlerKnow
    If TextField = "" Or TiaoJian = "" Then
        StrSql = "select * from View_AirportSzdmInfo where locale='" & Locale & "'"
    Else
        TextField = FiltQuoteChar(TextField)
        TextField = "%" & TextField & "%"
        StrSql = "select * from View_AirportSzdmInfo where " & TiaoJian & " like '" & TextField & "'" _
        & " and locale='" & Locale & "'"
    End If
    Set SearchAirportInfo = ObjDml.ExeSelect(StrSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchAirportInfo.EOF And SearchAirportInfo.BOF Then
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

'功能:根据条件查询视图View_SzdmCityName信息记录集
'返回值:组信息记录集
'输入参数：TextFileld:字段  TiaoJian:条件值
Public Function SearchSzdmCityInfo(ByVal TextField As String, ByVal TiaoJian As String, ByVal Locale As String) As Recordset
    Dim ObjDml As Object
    Dim StrSql
    Dim ErrNum As Long
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    On Error GoTo ErrorHandlerKnow
    If TextField = "" Or TiaoJian = "" Then
        StrSql = "select * from View_SzdmCityName where locale='" & Locale & "'"
    Else
        TextField = FiltQuoteChar(TextField)
        TextField = "%" & TextField & "%"
        StrSql = "select * from View_SzdmCityName where " & TiaoJian & " like '" & TextField & "'" _
                & " and locale='" & Locale & "'"
    End If
    Set SearchSzdmCityInfo = ObjDml.ExeSelect(StrSql, DbClass)
    On Error GoTo ErrorHandlerUnKnow
    If SearchSzdmCityInfo.EOF And SearchSzdmCityInfo.BOF Then
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
'功能：添加CityLocale表
'返回值：无
'输入参数:
'       CityID
'       CityName
'       Locale
Public Function AddCityLocale(ByVal CityID As String, ByVal CityName As String, _
                                 ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql, strSqlA As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As New ADODB.Recordset
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    CityName = FiltQuoteChar(CityName)
    CityID = FiltQuoteChar(CityID)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Insert into CityLocale (CityID,CityName,Locale) values " _
            & "('" & CityID & "','" & CityName & "','" & Locale & "')"
    strSqlA = "select * from CityLocale where locale='" & Locale & "' and Cityid='" & CityID & "'"
    '执行插入CityLocale表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(strSqlA, DbClass)
    If rs.EOF And rs.BOF Then
    Ret = ObjDml.ExeInsert(StrSql, DbClass)
    Set ObjDml = Nothing
    Else
    GoTo ErrorHandler
    End If
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能：添加City,Cityphone,Citylocale
'返回值：CityID
'输入参数:
'       Cityinfo    (6)
'       CityID      (0)
'       CityName    (1)
'       CityDepict  (2)
'       ProvinceID  (3)
'       PyCity      (4)
'       CityPhone   (5)
'       Locale      (6)
Public Function AddCityInfo(ByRef CityInfo) As String
    Dim ObjDml As Object
    Dim StrSql, strSqlA, strSqlB As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    CityInfo(0) = FiltQuoteChar(CStr(CityInfo(0)))
    CityInfo(1) = FiltQuoteChar(CStr(CityInfo(1)))
    CityInfo(2) = FiltQuoteChar(CStr(CityInfo(2)))
    CityInfo(3) = FiltQuoteChar(CStr(CityInfo(3)))
    CityInfo(4) = FiltQuoteChar(CStr(CityInfo(4)))
    CityInfo(5) = FiltQuoteChar(CStr(CityInfo(5)))
    
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "select * from CityLocale where cityid='" & CityInfo(0) & "' and cityname='" & CityInfo(1) & "'"
    strSqlA = "Insert into City (CityID,CityDepict,ProvinceID,PyCity) values" _
            & "('" & CityInfo(0) & "','" & CityInfo(2) & "'," _
            & "'" & CityInfo(3) & "','" & CityInfo(4) & "')"
    strSqlB = "Insert into CityPhone (CityID,CityPhone) values" _
             & "('" & CityInfo(0) & "','" & CityInfo(5) & "')"
    '执行插入LoginInfo表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(StrSql, DbClass)
    If rs.EOF And rs.BOF Then
    Ret = ObjDml.ExeInsert(strSqlA, DbClass)
    Ret = ObjDml.ExeInsert(strSqlB, DbClass)
    Ret = AddCityLocale(CityInfo(0), CityInfo(1), CityInfo(6))
    AddCityInfo = CityInfo(0)
    Set ObjDml = Nothing
    Else
    GoTo ErrorHandler
    End If
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能：添加SzdmCity,UseSzdmCity
'返回值：无
'输入参数:
'       SzdmCity  (4)
'       CsSzdm    (0)
'       CityID    (1)
'       UseObject (2)
'       Area      (3)
'       SendAllow (4)
Public Function AddSzdmCity(ByRef SzdmCity)
    Dim ObjDml As Object
    Dim StrSql, strSqlA, strSqlB As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As Recordset
    
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    SzdmCity(0) = FiltQuoteChar(CStr(SzdmCity(0)))
    SzdmCity(1) = FiltQuoteChar(CStr(SzdmCity(1)))
   
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "select * from SzdmCity where CsSzdm='" & SzdmCity(0) & "' and CityID='" & SzdmCity(1) & "'"
    strSqlA = "Insert into SzdmCity (CityID,CsSzdm) values" _
            & "('" & SzdmCity(1) & "','" & SzdmCity(0) & "')"
    strSqlB = "Insert into UseSzdmCity (CsSzdm,UseObject,Area,SendAllow) values" _
             & "('" & SzdmCity(0) & "','" & SzdmCity(2) & "','" & SzdmCity(3) & "','" & SzdmCity(4) & "')"
    '执行插入SzdmCity,UseSzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(StrSql, DbClass)
    If rs.EOF And rs.BOF Then
    Ret = ObjDml.ExeInsert(strSqlA, DbClass)
    Ret = ObjDml.ExeInsert(strSqlB, DbClass)
    Set ObjDml = Nothing
    Else
    GoTo ErrorHandler
    End If
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能：添加UseSzdmCity
'返回值：无
'输入参数:SzdmCity  (3)
'       CsSzdm    (0)
'       UseObject (1)
'       Area      (2)
'       SendAllow (3)
Public Function AddSzdmUseObject(ByRef SzdmUseObject)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As Recordset
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Insert into UseSzdmCity (CsSzdm,UseObject,Area,SendAllow) values" _
             & "('" & SzdmUseObject(0) & "','" & SzdmUseObject(1) & "'," _
             & "'" & SzdmUseObject(2) & "','" & SzdmUseObject(3) & "')"
    '执行插入UseSzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeInsert(StrSql, DbClass)
    Set ObjDml = Nothing
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能：修改CityLocale表信息
'返回值：无
'输入参数:CityID
'       Locale
'       CityName
Public Function EditCityLocale(ByVal CityID As String, ByVal CityName As String, _
                                 ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    CityName = FiltQuoteChar(CityName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Update CityLocale set CityName='" & CityName & "'" _
            & " where CityId='" & CityID & "' and locale='" & Locale & "'"
    '执行修改CityLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(StrSql, DbClass)
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
'功能：修改CityLocale表信息
'返回值：无
'输入参数:CityID
'       Locale
'       CityName
Public Function EditAirPortLocale(ByVal JcSzdm As String, ByVal AirportName As String, _
                                 ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    AirportName = FiltQuoteChar(AirportName)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Update AirportLocale set AirportName='" & AirportName & "'" _
            & " where jcszdm='" & JcSzdm & "' and locale='" & Locale & "'"
    '执行修改AirportLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(StrSql, DbClass)
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

'功能：修改City表信息
'返回值：无
'输入参数:CityID
'       CityDepict
'       ProvinceID
'       pyCity
Public Function EditCity(ByVal CityID As String, ByVal CityDepict As String, _
                                 ByVal ProvinceID As String, ByVal pyCity As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    CityDepict = FiltQuoteChar(CityDepict)
    ProvinceID = FiltQuoteChar(ProvinceID)
    pyCity = FiltQuoteChar(pyCity)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Update City set CityDepict='" & CityDepict & "'," _
            & "ProvinceID='" & ProvinceID & "'," & " pyCity ='" & pyCity & "'" _
            & " where CityId='" & CityID & "'"
    '执行修改City表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(StrSql, DbClass)
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

'功能：修改UseSzdmCity表信息
'返回值：无
'输入参数:CsSzdm
'       UseObject
'       Area
'       SendAllow
Public Function EditUseSzdmCity(ByVal CsSzdm As String, ByVal UseObject As String, _
                                 ByVal Area As String, ByVal SendAllow As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    Area = FiltQuoteChar(Area)
    SendAllow = FiltQuoteChar(SendAllow)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Update UseSzdmCity set Area='" & Area & "'," _
            & "SendAllow='" & SendAllow & "'" _
            & " where CsSzdm='" & CsSzdm & "'  and UseObject='" & UseObject & "'"
    '执行修改UseSzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(StrSql, DbClass)
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
'功能：删除CityLocale表信息
'返回值：无
'输入参数:CityID,Locale
Public Function DelCityLocale(ByVal CityID As String, ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "delete from CityLocale where CityID ='" & CityID & "'" _
                & " and Locale='" & Locale & "'"
    '执行删除CityLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(StrSql, DbClass)
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

'功能：删除UseSzdmCity表信息
'返回值：无
'输入参数:CsSzdm,UseObject
Public Function DelUseSzdmCity(ByVal CsSzdm As String, ByVal UseObject As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "delete from UseSzdmCity where CsSzdm ='" & CsSzdm & "'" _
                & " and UseObject='" & UseObject & "'"
    '执行删除UseSzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(StrSql, DbClass)
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
'功能：删除AirPortLocale表信息
'返回值：无
'输入参数:JcSzdm,Locale
Public Function DelAirPortLocale(ByVal JcSzdm As String, ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "delete from AirportLocale where JcSzdm ='" & JcSzdm & "'" _
                & " and Locale='" & Locale & "'"
    '执行删除AirportLocale表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(StrSql, DbClass)
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
'功能：删除JcSzdmAirPort,AirPortLocale表信息
'返回值：无
'输入参数:JcSzdm
Public Function DelSzdmAirPort(ByVal JcSzdm As String)
    Dim ObjDml As Object
    Dim strSqlA, strSqlB As String
    Dim Ret As Long
    Dim ErrNum As Long
    
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSqlA = "delete from AirportLocale where JcSzdm ='" & JcSzdm & "'"
    strSqlB = "delete from SzdmAirPort where JcSzdm ='" & JcSzdm & "'"
    '执行删除AirportLocale,SzdmAirPort表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeDelete(strSqlA, DbClass)
    Ret = ObjDml.ExeDelete(strSqlB, DbClass)
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

'功能：删除UseSzdmCity表,View_SzdmCity视图信息
'返回值：无
'输入参数:CsSzdm
Public Function DelSzdmCity(ByVal CsSzdm As String)
    Dim ObjDml As Object
    Dim StrSql, strSqlA, strSqlB, strSqlC, strSqlD As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "select * from SzdmAirport where CsSzdm='" & CsSzdm & "'"
    strSqlA = "delete from UseSzdmCity where CsSzdm ='" & CsSzdm & "'"
    strSqlB = "delete from SzdmAirport where CsSzdm ='" & CsSzdm & "'"
    strSqlD = "delete from SzdmCity where CsSzdm='" & CsSzdm & "'"
    '执行删除UseSzdmCity,View_SzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(StrSql, DbClass)
    If Not rs.EOF And Not rs.BOF Then
    For i = 1 To rs.RecordCount
    strSqlC = "delete from AirportLocale where JcSzdm='" & rs("JcSzdm") & "'"
    Ret = ObjDml.ExeDelete(strSqlC, DbClass)
    rs.MoveNext
    Next
    End If
    Ret = ObjDml.ExeDelete(strSqlA, DbClass)
    Ret = ObjDml.ExeDelete(strSqlB, DbClass)
    Ret = ObjDml.ExeDelete(strSqlD, DbClass)
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
'功能：删除Cityphone,CityLocale表信息
'返回值：无
'输入参数:CityID
Public Function DelCity(ByVal CityID As String)
    Dim ObjDml As Object
    Dim StrSql, strSqlA, strSqlB, strSqlC As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim CsSzdm As String
    Dim rs As New ADODB.Recordset
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSqlA = "delete from CityLocale where CityID ='" & CityID & "'"
    strSqlB = "delete from CityPhone where CityID ='" & CityID & "'"
    strSqlC = "delete from city where cityid='" & CityID & "'"
    StrSql = "select * from SzdmCity where CityID='" & CityID & "'"
    '执行删除Cityphone,CityLocale表,View_City视图操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(StrSql, DbClass)
    If Not rs.EOF And Not rs.BOF Then
    CsSzdm = rs("CsSzdm")
    Ret = DelSzdmCity(CsSzdm)
    End If
    Ret = ObjDml.ExeDelete(strSqlA, DbClass)
    Ret = ObjDml.ExeDelete(strSqlB, DbClass)
    Ret = ObjDml.ExeDelete(strSqlC, DbClass)
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
'功能：修改SzdmAirport表信息
'返回值：无
'输入参数:JcSzdm
'         AirportDepict
Public Function EditSzdmAirport(ByVal JcSzdm As String, ByVal AirportDepict As String)
    Dim ObjDml As Object
    Dim StrSql As String
    Dim Ret As Long
    Dim ErrNum As Long
  
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    AirportDepict = FiltQuoteChar(AirportDepict)
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Update SzdmAirport set AirportDepict='" & AirportDepict & "'" _
                        & " where JcSzdm='" & JcSzdm & "'"
    '执行修改SzdmAirport表操作
    On Error GoTo ErrorHandlerKnow
    Ret = ObjDml.ExeUpdate(StrSql, DbClass)
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
'功能：添加SzdmAirport,AirPortLocale
'返回值：无
'输入参数:
'       SzdmAirport      (4)
'       JcSzdm           (0)
'       AirPortDepict    (1)
'       CsSzdm           (2)
'       Locale           (3)
'       AirportName      (4)
Public Function AddSzdmAirport(ByRef SzdmAirPort) As String
    Dim ObjDml As Object
    Dim StrSql, strSqlA As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As New Recordset
    ErrorNo = 0
    On Error GoTo ErrorHandlerKnow
    SzdmAirPort(0) = FiltQuoteChar(CStr(SzdmAirPort(0)))
    SzdmAirPort(1) = FiltQuoteChar(CStr(SzdmAirPort(1)))
    SzdmAirPort(2) = FiltQuoteChar(CStr(SzdmAirPort(2)))
    SzdmAirPort(4) = FiltQuoteChar(CStr(SzdmAirPort(4)))
    '创建DML对象
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    strSqlA = "select * from SzdmAirport where JcSzdm='" & SzdmAirPort(0) & "'"
    StrSql = "Insert into SzdmAirport (JcSzdm,AirportDepict,CsSzdm) values" _
            & "('" & SzdmAirPort(0) & "','" & SzdmAirPort(1) & "','" & SzdmAirPort(2) & "')"
    
    '执行插入SzdmCity,UseSzdmCity表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(strSqlA, DbClass)
    If rs.EOF And rs.BOF Then
    Ret = ObjDml.ExeInsert(StrSql, DbClass)
    Ret = AddAirportLocale(SzdmAirPort(0), SzdmAirPort(4), SzdmAirPort(3))
    AddSzdmAirport = SzdmAirPort(0)
    Set ObjDml = Nothing
    Else
    GoTo ErrorHandler
    End If
    GetObjectContext.SetComplete
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10683
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能：添加AirportLocale表
'返回值：无
'输入参数:
'       JcSzdmID
'       AirportName
'       Locale
Public Function AddAirportLocale(ByVal JcSzdm As String, ByVal AirportName As String, _
                                 ByVal Locale As String)
    Dim ObjDml As Object
    Dim StrSql, strSqlA As String
    Dim Ret As Long
    Dim ErrNum As Long
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    ErrorNo = 0
    '创建DML对象
    On Error GoTo ErrorHandlerKnow
    AirportName = FiltQuoteChar(AirportName)
    JcSzdm = FiltQuoteChar(JcSzdm)
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_DML.ClsDML")
    StrSql = "Insert into AirportLocale (JcSzdm,AirportName,Locale) values " _
            & "('" & JcSzdm & "','" & AirportName & "','" & Locale & "')"
    strSqlA = "select * from AirportLocale where JcSzdm='" & JcSzdm & "' and Locale='" & Locale & "'"
    '执行插入AirportLocale表操作
    On Error GoTo ErrorHandlerKnow
    Set rs = ObjDml.ExeSelect(strSqlA, DbClass)
    If rs.EOF And rs.BOF Then
    Ret = ObjDml.ExeInsert(StrSql, DbClass)
    Set ObjDml = Nothing
    Else
    GoTo ErrorHandler
    End If
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    ErrorNo = 10792
    Err.Clear
ErrorHandler:
    ErrorNo = 10790
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

'功能:根据CityID获取指定视图View_CityInfo信息记录集
'返回值:信息记录集
'输入参数：CityID,Locale
Public Function GetCityInfo(ByVal CityID As String, Optional ByVal Locale As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        StrSql = "select * from View_CityInfo where CityID='" & CityID & "'"
                    
    Else
        StrSql = "select * from View_CityInfo where CityID='" & CityID & "'" _
                    & " and locale='" & Locale & "'"
    End If
    
    On Error GoTo ErrorHandlerKnow
    Set GetCityInfo = ObjDml.ExeSelect(StrSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetCityInfo.EOF Then
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
'功能:根据CityID,Locale获取信息
'返回值:指定语言功能名
'输入参数：CityID,Locale
Public Function GetCityLocale(ByVal CityID As String, Optional ByVal Locale As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
    StrSql = "select * from CityLocale where CityID='" & CityID & "'"
    
    Else
    StrSql = "select * from CityLocale where CityID='" & CityID & "'" _
                & " and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set GetCityLocale = ObjDml.ExeSelect(StrSql, DbClass)
    
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
'功能:根据JcSzdm,Locale获取信息
'返回值:指定语言功能名
'输入参数：JcSzdm,Locale
Public Function GetAirPortLocale(ByVal JcSzdm As String, Optional ByVal Locale As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If JcSzdm = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
    StrSql = "select * from AirportLocale where JcSzdm='" & JcSzdm & "'"
    
    Else
    StrSql = "select * from AirportLocale where JcSzdm='" & JcSzdm & "'" _
                & " and Locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set GetAirPortLocale = ObjDml.ExeSelect(StrSql, DbClass)
    
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
'功能:根据CsSzdm,Useobject获取信息
'输入参数：CsSzdm,Useobject
Public Function GetUseSzdmCity(ByVal CsSzdm As String, Optional ByVal UseObject As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CsSzdm = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If UseObject <> "" Then
        StrSql = "select * from UseSzdmCity where CsSzdm='" & CsSzdm & "'" _
                & " and UseObject='" & UseObject & "'"
    Else
        StrSql = "select * from UseSzdmCity where CsSzdm='" & CsSzdm & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set GetUseSzdmCity = ObjDml.ExeSelect(StrSql, DbClass)
    
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

'功能:根据CityID获取City信息记录集
'返回值:City信息记录集
'输入参数：CityID
Public Function GetCity(ByVal CityID As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from City where CityID='" & CityID & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetCity = ObjDml.ExeSelect(StrSql, DbClass)
    
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
'功能:根据JcSzdm获取SzdmAirport信息记录集
'返回值:SzdmAirport信息记录集
'输入参数：JcSzdm
Public Function GetSzdmAirport(ByVal JcSzdm As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If JcSzdm = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from SzdmAirport where JcSzdm='" & JcSzdm & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetSzdmAirport = ObjDml.ExeSelect(StrSql, DbClass)
    
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

'功能:根据CityID获取CityPhone信息记录集
'返回值:Cityphone信息记录集
'输入参数：CityID
Public Function GetCityPhone(ByVal CityID As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from CityPhone where CityID='" & CityID & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetCityPhone = ObjDml.ExeSelect(StrSql, DbClass)
    
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
'功能:获取Locale信息记录集
'返回值:Locale信息记录集
Public Function GetLocale() As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from Locale "
    
    On Error GoTo ErrorHandlerKnow
    Set GetLocale = ObjDml.ExeSelect(StrSql, DbClass)
    
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

'功能:取用户对象记录集
'返回值:UseObject记录集
'输入参数：无
Public Function GetUseObject() As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long

    ErrorNo = 0

    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from UseObject"

    On Error GoTo ErrorHandlerKnow
    Set GetUseObject = ObjDml.ExeSelect(StrSql, DbClass)

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

'功能:根据CityID获取SzdmCity信息记录集
'返回值:SzdmCity信息记录集
'输入参数：CityID
Public Function GetSzdmCity(ByVal CityID As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from SzdmCity where CityID='" & CityID & "'"
    
    On Error GoTo ErrorHandlerKnow
    Set GetSzdmCity = ObjDml.ExeSelect(StrSql, DbClass)
    
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

'功能:根据CityID获取指定视图View_SzdmCityInfo信息记录集
'返回值:信息记录集
'输入参数：CityID
Public Function GetSzdmCityInfo(ByVal CityID As String, ByVal Locale As String, ByVal UseObject As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If CityID = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    StrSql = "select * from View_SzdmCityInfo where CityID='" & CityID & "'" _
            & " and  UseObject='" & UseObject & "'" & " and locale='" & Locale & "'"
                
    On Error GoTo ErrorHandlerKnow
    Set GetSzdmCityInfo = ObjDml.ExeSelect(StrSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetSzdmCityInfo.EOF Then
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

'功能:根据JcSzdm获取指定视图View_AirportSzdmInfo信息记录集
'返回值:信息记录集
'输入参数：JcSzdm,Locale
Public Function GetAirportInfo(ByVal JcSzdm As String, Optional ByVal Locale As String) As Recordset
    Dim StrSql As String
    Dim ObjDml As Object
    Dim ErrNum As Long
    
    ErrorNo = 0
    If JcSzdm = "" Then
        ErrorNo = 10793
        GoTo ErrorHandlerKnow
    End If
    On Error GoTo ErrorHandlerUnKnow
    Set ObjDml = CreateObject("Com_Dml.ClsDml")
    If Locale = "" Then
        StrSql = "select * from View_AirportSzdmInfo where JcSzdm='" & JcSzdm & "'"
                    
    Else
        StrSql = "select * from View_AirportSzdmInfo where JcSzdm='" & JcSzdm & "'" _
                    & " and locale='" & Locale & "'"
    End If
    On Error GoTo ErrorHandlerKnow
    Set GetAirportInfo = ObjDml.ExeSelect(StrSql, DbClass)
    
    On Error GoTo ErrorHandlerUnKnow
    If GetAirportInfo.EOF Then
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

