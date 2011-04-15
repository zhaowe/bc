VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Private Const DbClass = 91    '数据库指向SZ_Maintain
Private ErrorNo As Long
'报表2，根据DeviceClass,Person,Status求数量和
Public Function GetCount(ByVal ThisDbClass As Integer, ByVal DeviceClass As String, Optional Person As String, Optional ByVal Status As String, Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Integer
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    
    If DeviceClass = "All" Then
       If Person = "All" Then
          Select Case Status
          Case "All"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "')"
          Case "Dealed"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       Else
          Select Case Status
          Case "All"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "'"
          Case "Dealed"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       End If
    Else
       If Person = "All" Then
          Select Case Status
          Case "All"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "'"
          Case "Dealed"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       Else
          Select Case Status
          Case "All"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "'"
          Case "Dealed"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       End If
    End If
    
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("sz_maintain_Dml.sz_maintain_ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDML.ExeSelect(Str, ThisDbClass)
    GetCount = Rs("CountNum")
    On Error GoTo ErrorHandlerUnknow
    Set Rs = Nothing
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    '未知错误
    ErrorNo = 20010
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnknow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDML Is Nothing Then
        Set ObjDML = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description

End Function

'根据操作员输入的简单信息，获取数据库的相关记录集；模糊查询。
Public Function GetReportRs(ByVal ThisDbClass As Integer, ByVal Status As String, Optional BeginDate As String, Optional EndDate As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Select Case Status
'    If Status = "All" Then
    Case "All"
       Str = "Select * From MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') order by SXRQ,ID"
'    Else
    Case "Dealed"
       Str = "Select * From MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') order by SXRQ,ID"
    Case Else
    End Select
'    End If
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("sz_maintain_Dml.sz_maintain_ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set GetReportRs = ObjDML.ExeSelect(Str, ThisDbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    '未知错误
    ErrorNo = 20011
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnknow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDML Is Nothing Then
        Set ObjDML = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description

End Function

'报表3，根据DeviceClass,Person,Status求数量和
Public Function GetModeCount(ByVal ThisDbClass As Integer, ByVal DeviceClass As String, Optional Person As String, Optional ByVal Status As String, Optional ByVal Mode As String, Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Integer
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    
    Select Case Mode
    Case "自修"
       If DeviceClass = "All" Then
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and (WXFS='自修' or WXFS is null)"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and (WXFS='自修' or WXFS is null)"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and (WXFS='自修' or WXFS is null)"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and (WXFS='自修' or WXFS is null)"
             Case Else
             End Select
          End If
       Else
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and (WXFS='自修' or WXFS is null)"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and (WXFS='自修' or WXFS is null)"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and (WXFS='自修' or WXFS is null)"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and (WXFS='自修' or WXFS is null)"
             Case Else
             End Select
          End If
       End If
    Case "保修"
       If DeviceClass = "All" Then
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXFS='保修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='保修'"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and WXFS='保修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='保修'"
             Case Else
             End Select
          End If
       Else
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXFS='保修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='保修'"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and WXFS='保修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='保修'"
             Case Else
             End Select
          End If
       End If
    Case "外修"
       If DeviceClass = "All" Then
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXFS='外修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='外修'"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and WXFS='外修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='外修'"
             Case Else
             End Select
          End If
       Else
          If Person = "All" Then
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXFS='外修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='外修'"
             Case Else
             End Select
          Else
             Select Case Status
             Case "All"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and WXFS='外修'"
             Case "Dealed"
                  Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单') and WXFS='外修'"
             Case Else
             End Select
          End If
       End If
    End Select
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("sz_maintain_Dml.sz_maintain_ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDML.ExeSelect(Str, ThisDbClass)
    GetModeCount = Rs("CountNum")
    On Error GoTo ErrorHandlerUnknow
    Set Rs = Nothing
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    '未知错误
    ErrorNo = 20012
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnknow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDML Is Nothing Then
        Set ObjDML = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description

End Function

'报表4，根据DeviceClass,Person,DateType求数量和
Public Function GetAbandonCount(ByVal ThisDbClass As Integer, ByVal DeviceClass As String, Optional Person As String, Optional ByVal DateType As String, Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Integer
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    
    If DeviceClass = "All" Then
       If Person = "All" Then
          Select Case DateType
          Case "SXRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBZT='已报废'"
          Case "BFRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (QJRQ between '" & BeginDate & "' and '" & EndDate & "') and SBZT='已报废'"
          Case Else
          End Select
       Else
          Select Case DateType
          Case "SXRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and SBZT='已报废'"
          Case "BFRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (QJRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and SBZT='已报废'"
          Case Else
          End Select
       End If
    Else
       If Person = "All" Then
          Select Case DateType
          Case "SXRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and SBZT='已报废'"
          Case "BFRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (QJRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and SBZT='已报废'"
          Case Else
          End Select
       Else
          Select Case DateType
          Case "SXRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and SBZT='已报废'"
          Case "BFRQ"
               Str = "select isnull(sum(SL),0) as CountNum from MaintainData where (QJRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and SBZT='已报废'"
          Case Else
          End Select
       End If
    End If
    
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("sz_maintain_Dml.sz_maintain_ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDML.ExeSelect(Str, ThisDbClass)
    GetAbandonCount = Rs("CountNum")
    On Error GoTo ErrorHandlerUnknow
    Set Rs = Nothing
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    '未知错误
    ErrorNo = 20013
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnknow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDML Is Nothing Then
        Set ObjDML = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description

End Function

'报表2，根据DeviceClass,Person,Status求数量和
Public Function GetFare(ByVal ThisDbClass As Integer, ByVal DeviceClass As String, Optional Person As String, Optional ByVal Status As String, Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Currency
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    
    If DeviceClass = "All" Then
       If Person = "All" Then
          Select Case Status
          Case "All"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "')"
          Case "Dealed"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       Else
          Select Case Status
          Case "All"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "'"
          Case "Dealed"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       End If
    Else
       If Person = "All" Then
          Select Case Status
          Case "All"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "'"
          Case "Dealed"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       Else
          Select Case Status
          Case "All"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "'"
          Case "Dealed"
               Str = "select isnull(sum(WXF*SL),0) as FareNum from MaintainData where (XHRQ between '" & BeginDate & "' and '" & EndDate & "') and SBLX='" & DeviceClass & "' and WXRY='" & Person & "' and ZT in ('已取走','已入库','已发货','已报废','已修好','销单')"
          Case Else
          End Select
       End If
    End If
    
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("sz_maintain_Dml.sz_maintain_ClsDml")
    On Error GoTo ErrorHandlerKnow
    Set Rs = ObjDML.ExeSelect(Str, ThisDbClass)
    GetFare = Rs("FareNum")
    On Error GoTo ErrorHandlerUnknow
    Set Rs = Nothing
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    '未知错误
    ErrorNo = 20010
    Err.Clear
ErrorHandlerKnow:
    '已知错误
    If Err.Number <> 0 Then
        If Err.Number < 10000 Then
            GoTo ErrorHandlerUnknow
        Else
            ErrNum = Err.Number
        End If
    Else
        ErrNum = ErrorNo
    End If
    If Not ObjDML Is Nothing Then
        Set ObjDML = Nothing
    End If
    On Error GoTo EndError
    If ErrNum <> 0 Then
        Err.Raise ErrNum
    End If
    Exit Function
EndError:
    Err.Raise ErrNum, Err.Source, Err.Description

End Function


