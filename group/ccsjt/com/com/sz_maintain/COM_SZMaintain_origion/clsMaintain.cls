VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Private Const DbClass = 3     '���ݿ�ָ��SZ_Maintain
Private ErrorNo As Long

'��ü�¼����,DataType������
Public Function GetRs(ByVal RsType As String, Optional ByVal Field As String, Optional ByVal Value As String = "", Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Select Case RsType
    Case "Account"
         If Value <> "" Then
            Str = "Select * From Account Where " & Field & "='" & Value & "' order by Account"
         Else
            Str = "Select * From Account order by Account"
         End If
    Case "MaintainPerson"
         Str = "Select * From Account Where IsMaintain=1 order by Account"
    Case "DeviceClass"
         Str = "Select * From DeviceClass order by Attribute desc,DeviceClass"
    Case "Status"
         Str = "Select * From Status order by StatusID"
    Case "MaintainData"
         If BeginDate = "" And EndDate = "" Then
            If Field = "" Then
               If Value = "" Then
                  Str = "Select * From View_Maintain order by ID desc"
               Else
                  Str = "Select * From View_Maintain where (ID like '%" & Value & "%') or (SXDW like '%" & Value & "%') or (SBLX like '%" & Value & "%') or (JQXH like '%" & Value & "%') or (SL like '%" & Value & "%') or (SBBH like '%" & Value & "%') or (XLH like '%" & Value & "%') or (LXRY like '%" & Value & "%') or (LXDH like '%" & Value & "%') or (KHDZ like '%" & Value & "%')" & _
                        " or (SBPH like '%" & Value & "%') or (FZ like '%" & Value & "%') or (SBZT like '%" & Value & "%') or (WXFS like '%" & Value & "%') or (GZXX like '%" & Value & "%') or (SXRK like '%" & Value & "%') or (WXRY like '%" & Value & "%') or (ZT like '%" & Value & "%') or (GHBJ like '%" & Value & "%') or (BZ like '%" & Value & "%') or (QJRY like '%" & Value & "%') or (GZPD like '%" & Value & "%') or (SHBJ like '%" & Value & "%') or (BFLY like '%" & Value & "%') order by ID desc"
               End If
            Else '�豸����
                  Str = "Select * From View_Maintain where SBPH='" & Value & "' order by ID desc"
            End If
         Else
            If Value = "" Then
               Str = "Select * From View_Maintain where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') or (XHRQ between '" & BeginDate & "' and '" & EndDate & "') or (QJRQ between '" & BeginDate & "' and '" & EndDate & "') order by ID desc"
            Else
               Str = "Select * From View_Maintain where ((ID like '%" & Value & "%') or (SXDW like '%" & Value & "%') or (SBLX like '%" & Value & "%') or (JQXH like '%" & Value & "%') or (SL like '%" & Value & "%') or (SBBH like '%" & Value & "%') or (XLH like '%" & Value & "%') or (LXRY like '%" & Value & "%') or (LXDH like '%" & Value & "%') or (KHDZ like '%" & Value & "%')" & _
                     " or (SBPH like '%" & Value & "%') or (FZ like '%" & Value & "%') or (SBZT like '%" & Value & "%') or (WXFS like '%" & Value & "%') or (GZXX like '%" & Value & "%') or (SXRK like '%" & Value & "%') or (WXRY like '%" & Value & "%') or (ZT like '%" & Value & "%') or (GHBJ like '%" & Value & "%') or (BZ like '%" & Value & "%') or (QJRY like '%" & Value & "%')) or (GZPD like '%" & Value & "%') or (SHBJ like '%" & Value & "%') or (BFLY like '%" & Value & "%')" & _
                     " and ((SXRQ between '" & BeginDate & "' and '" & EndDate & "') or (XHRQ between '" & BeginDate & "' and '" & EndDate & "') or (QJRQ between '" & BeginDate & "' and '" & EndDate & "')) order by ID desc"
            End If
         End If

    Case Else
              
    
    End Select
    
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetRs = ObjDML.ExeSelect(Str, DbClass)
'    Set GetRs = clsDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20000
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'��ȡĬ�ϼ�¼����
Public Function GetDefaultQuery(ByVal Grade As String, ByVal WXRY As String, ByVal GroupClass As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Select Case Grade
    Case "Administrator"
         Str = "Select * From View_Maintain Where SBZT not in ('��ȡ��','�����','�ѷ���','�ѱ���','����') order by ID desc"
    Case "Group"
'         Str = "Select * From View_Maintain Where (WXXZ='" & GroupClass & "' or WXRY='" & WXRY & "' or SBLX='����') and SBZT not in ('��ȡ��','�����','�ѷ���','�ѱ���','����') and ZT not in ('���޺�','������') order by ID desc"
         Str = "Select * From View_Maintain Where (WXXZ='" & GroupClass & "' or WXRY='" & WXRY & "' or SBLX='����') and ZT not in ('���޺�','������','��ȡ��','�����','�ѷ���','�ѱ���','����') order by ID desc"
    Case "Member"
'         Str = "Select * From View_Maintain Where WXRY='" & WXRY & "' and SBZT not in ('��ȡ��','�����','�ѷ���','�ѱ���','����') and ZT not in ('���޺�','������',) order by ID desc"
         Str = "Select * From View_Maintain Where WXRY='" & WXRY & "' and ZT not in ('���޺�','������','��ȡ��','�����','�ѷ���','�ѱ���','����') order by ID desc"
    Case "Operator"   'Operator��Administrator��ͬ
         Str = "Select * From View_Maintain Where SBZT not in ('��ȡ��','�����','�ѷ���','�ѱ���','����') order by ID desc"
    Case "Guest"   'Guest��Administrator��ͬ
         Str = "Select * From View_Maintain Where SBZT not in ('��ȡ��','�����','�ѷ���','�ѱ���','����') order by ID desc"
    Case Else
    
    End Select


    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetDefaultQuery = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20001
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
'���ݲ���Ա����ļ���Ϣ����ȡ���ݿ����ؼ�¼����ģ����ѯ��
Public Function GetMaintainRs(ByVal SXDW As String, Optional LXRY As String = "") As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    If LXRY = "" Then
       Str = "Select * From MaintainData where SXDW like '%" & SXDW & "%' order by ID desc"
    Else
       Str = "Select * From MaintainData where LXRY like '%" & LXRY & "%' order by ID desc"
      
    End If
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetMaintainRs = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20002
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
Public Function GetIDRs(ByVal ID As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Select * From View_Maintain where ID='" & ID & "'"
   
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetIDRs = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20003
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'�����ʺ�
Public Function UpdateAccount(ByVal Privilege As String, ByVal Account As String, ByVal Password As String, Optional ByVal Grade As String, Optional ByVal Name As String, Optional ByVal GroupClass As String, Optional ByVal IsMaintain As String, Optional ByVal Description As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    If Privilege = "Administrator" Then
       Str = "Update Account set Account='" & Account & "',Password='" & Password & "',Grade='" & Grade & "',Name='" & Name & "',GroupClass='" & GroupClass & "',IsMaintain='" & IsMaintain & "',Description='" & Description & "' where Account='" & Account & "'"
    Else
       Str = "Update Account set Password='" & Password & "' Where Account='" & Account & "'"
    End If
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    UpdateAccount = ObjDML.ExeUpdate(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20004
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
'ɾ����¼
Public Function DeleteRs(ByVal RsType As String, ByVal ID As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    
    Select Case RsType
    Case "Account"
       Str = "delete from Account where Account='" & ID & "'"
    Case "MaintainData"
       Str = "delete from MaintainData where ID='" & ID & "'"
    Case "DeviceClass"
       Str = "delete from DeviceClass where DeviceClass='" & ID & "'"
    End Select
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    DeleteRs = ObjDML.ExeDelete(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20005
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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


'�����ʺ�
Public Function AddAccount(ByVal Account As String, ByVal Password As String, Optional ByVal Grade As String, Optional ByVal Name As String, Optional ByVal GroupClass As String, Optional ByVal IsMaintain As String, Optional ByVal Description As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Insert Into Account (Account,Password,Grade,Name,GroupClass,IsMaintain,Description) Values ('" & Account & "','" & Password & "','" & Grade & "','" & Name & "','" & GroupClass & "','" & IsMaintain & "','" & Description & " ')"
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    AddAccount = ObjDML.ExeInsert(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20006
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'�����ʺ�
Public Function UpdateMaintain(ByVal Grade As String, ByVal ID As String, ByVal SXDW As String, ByVal SXRQ As String, ByVal SBLX As String, ByVal JQXH As String, ByVal SL As String, ByVal SBBH As String, ByVal XLH As String, ByVal LXRY As String, ByVal LXDH As String, ByVal GZXX As String, ByVal SXRK As String, ByVal SFJJ As String, _
    ByVal WXRY As String, ByVal ZT As String, ByVal XHRQ As String, ByVal WXFS As String, ByVal GHBJ As String, ByVal WXF As String, ByVal BZ As String, ByVal QJRY As String, ByVal QJRQ As String, ByVal SBPH As String, ByVal FZ As String, ByVal SBZT As String, ByVal KHDZ As String, ByVal GZPD As String, ByVal SHBJ As String, ByVal BFLY As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    Select Case Grade
    Case "Administrator"
       Str = "Update MaintainData set SXDW='" & SXDW & "',SXRQ='" & SXRQ & "',SBLX='" & SBLX & "',JQXH='" & JQXH & "',SL='" & SL & "',SBBH='" & SBBH & "',XLH='" & XLH & "',LXRY='" & LXRY & "',LXDH='" & LXDH & "',GZXX='" & GZXX & "',SXRK='" & SXRK & "',SFJJ='" & SFJJ & _
             "',WXRY='" & WXRY & "',ZT='" & ZT & "',XHRQ='" & XHRQ & "',WXFS='" & WXFS & "',GHBJ='" & GHBJ & "',WXF=convert(money,'" & WXF & "'),BZ='" & BZ & "',QJRY='" & QJRY & "',QJRQ='" & QJRQ & "',SBPH='" & SBPH & "',FZ='" & FZ & "',SBZT='" & SBZT & "',KHDZ='" & KHDZ & "',GZPD='" & GZPD & "',SHBJ='" & SHBJ & "',BFLY='" & BFLY & "' where ID='" & ID & "'"
    Case "Operator"
       Str = "Update MaintainData set SXDW='" & SXDW & "',SXRQ='" & SXRQ & "',SBLX='" & SBLX & "',JQXH='" & JQXH & "',SL='" & SL & "',SBBH='" & SBBH & "',XLH='" & XLH & "',LXRY='" & LXRY & "',LXDH='" & LXDH & "',GZXX='" & GZXX & "',SXRK='" & SXRK & "',SFJJ='" & SFJJ & _
             "',SBZT='" & SBZT & "',BZ='" & BZ & "',QJRY='" & QJRY & "',QJRQ='" & QJRQ & "',SBPH='" & SBPH & "',FZ='" & FZ & "',KHDZ='" & KHDZ & "' where ID='" & ID & "'"
    Case Else
       Str = "Update MaintainData set WXRY='" & WXRY & "',ZT='" & ZT & "',XHRQ='" & XHRQ & "',WXFS='" & WXFS & "',GHBJ='" & GHBJ & "',WXF=convert(money,'" & WXF & "'),BZ='" & BZ & "',GZPD='" & GZPD & "',SHBJ='" & SHBJ & "',BFLY='" & BFLY & "' where ID='" & ID & "'"
             
    End Select
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    UpdateMaintain = ObjDML.ExeUpdate(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20007
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'�����ʺ�
Public Function AddMaintain(ByVal SXDW As String, ByVal SXRQ As String, ByVal SBLX As String, ByVal JQXH As String, ByVal SL As String, ByVal SBBH As String, ByVal XLH As String, ByVal LXRY As String, ByVal LXDH As String, ByVal GZXX As String, ByVal SXRK As String, ByVal SFJJ As String, ByVal SBPH As String, ByVal FZ As String, ByVal KHDZ As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Insert Into MaintainData (SXDW,SXRQ,SBLX,JQXH,SL,SBBH,XLH,LXRY,LXDH,GZXX,SXRK,SFJJ,SBPH,FZ,KHDZ,WXRY,ZT,WXFS,SBZT) Values ('" _
          & SXDW & "','" & SXRQ & "','" & SBLX & "','" & JQXH & "','" & SL & "','" & SBBH & "','" & XLH & "','" & LXRY & "','" & LXDH & "','" & GZXX & "','" & SXRK & "','" & SFJJ & "','" & SBPH & "','" & FZ & "','" & KHDZ & "','δָ��', '����','����','')"
  
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    AddMaintain = ObjDML.ExeInsert(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function
    
ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20008
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'Public Function AddLog(ByVal Account As String, ByVal Grade As String, ByVal OperaName As String, ByVal ID As String, ByVal SXDW As String, ByVal SXRQ As String, ByVal SBLX As String, ByVal JQXH As String, ByVal SL As String, ByVal SBBH As String, ByVal XLH As String, ByVal LXRY As String, ByVal LXDH As String, ByVal GZXX As String, ByVal SXRK As String, ByVal SFJJ As String, _
'                       ByVal WXRY As String, ByVal ZT As String, ByVal XHRQ As String, ByVal WXFS As String, ByVal GHBJ As String, ByVal WXF As String, ByVal BZ As String, ByVal QJRY As String, ByVal QJRQ As String, ByVal SBPH As String, ByVal FZ As String, ByVal SBZT As String) As Long
Public Function AddLog(ByVal Account As String, ByVal Grade As String, ByVal OperaName As String, ByVal ID As String) As Long
    
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    Dim Rs As Recordset
    
    ErrorNo = 0
    
'    Str = "Insert Into MaintainLog (Account,Grade,OperaName,ID,SXDW,SXRQ,SBLX,JQXH,SL,SBBH,XLH,LXRY,LXDH,GZXX,SXRK,SFJJ,WXRY,ZT,XHRQ,WXFS,GHBJ,WXF,BZ,QJRY,QJRQ,SBPH,FZ,SBZT) Values ('" _
'          & Account & "','" & Grade & "','" & OperaName & "','" & ID & "','" & SXDW & "','" & SXRQ & "','" & SBLX & "','" & JQXH & "','" & SL & "','" & SBBH & "','" & XLH & "','" & LXRY & "','" & LXDH & "','" & GZXX & "','" _
'          & SXRK & "','" & SFJJ & "','" & WXRY & "','" & ZT & "','" & XHRQ & "','" & WXFS & "','" & GHBJ & "',convert(money,'" & WXF & "'),'" & BZ & "','" & QJRY & "','" & QJRQ & "','" & SBPH & "','" & FZ & "','" & SBZT & "')"
 
    If OperaName = "ɾ��" Then
        Str = "Insert Into MaintainLog (Account,Grade,OperaName,ID) Values ('" & Account & " ','" & Grade & "','" & OperaName & "','" & ID & "')"
    Else
        Set Rs = GetIDRs(Trim(ID))
        'ID = Trim(request("ID"))
        SXDW = Trim(Rs("SXDW"))
        SXRQ = Trim(Rs("SXRQ"))
        SBLX = Trim(Rs("SBLX"))
        JQXH = Trim(Rs("JQXH"))
        SL = Trim(Rs("SL"))
        SBBH = Trim(Rs("SBBH"))
        XLH = Trim(Rs("XLH"))
        LXRY = Trim(Rs("LXRY"))
        LXDH = Trim(Rs("LXDH"))
        GZXX = Trim(Rs("GZXX"))
        SXRK = Trim(Rs("SXRK"))
        SFJJ = Trim(Rs("SFJJ"))
        If SFJJ = "False" Then
           SFJJ = "0"
        Else
           SFJJ = "1"
        End If
        WXRY = Trim(Rs("WXRY"))
        ZT = Trim(Rs("ZT"))
        XHRQ = Trim(Rs("XHRQ"))
        WXFS = Trim(Rs("WXFS"))
        GHBJ = Trim(Rs("GHBJ"))
        WXF = Trim(Rs("WXF"))
        BZ = Trim(Rs("BZ"))
        QJRY = Trim(Rs("QJRY"))
        QJRQ = Trim(Rs("QJRQ"))
        SBPH = Trim(Rs("SBPH"))
        FZ = Trim(Rs("FZ"))
        SBZT = Trim(Rs("SBZT"))
        KHDZ = Trim(Rs("KHDZ"))
        GZPD = Trim(Rs("GZPD"))
        SHBJ = Trim(Rs("SHBJ"))
        BFLY = Trim(Rs("BFLY"))
        Str = "Insert Into MaintainLog (Account,Grade,OperaName,ID,SXDW,SXRQ,SBLX,JQXH,SL,SBBH,XLH,LXRY,LXDH,GZXX,SXRK,SFJJ,WXRY,ZT,XHRQ,WXFS,GHBJ,WXF,BZ,QJRY,QJRQ,SBPH,FZ,SBZT,KHDZ,GZPD,SHBJ,BFLY) Values ('" _
              & Account & "','" & Grade & "','" & OperaName & "','" & ID & "','" & SXDW & "','" & SXRQ & "','" & SBLX & "','" & JQXH & "','" & SL & "','" & SBBH & "','" & XLH & "','" & LXRY & "','" & LXDH & "','" & GZXX & "','" _
              & SXRK & "','" & SFJJ & "','" & WXRY & "','" & ZT & "','" & XHRQ & "','" & WXFS & "','" & GHBJ & "',convert(money,'" & WXF & "'),'" & BZ & "','" & QJRY & "','" & QJRQ & "','" & SBPH & "','" & FZ & "','" & SBZT & "','" & KHDZ & "','" & GZPD & "','" & SHBJ & "','" & BFLY & "')"
    End If
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    AddLog = ObjDML.ExeInsert(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Set Rs = Nothing
    Exit Function
    
ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20009
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

Public Function GetLogRs(ByVal Value As String, Optional ByVal BeginDate As String, Optional ByVal EndDate As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
         If BeginDate = "" And EndDate = "" Then
'            If Field = "" Then
            If Value = "" Then
               Str = "Select * From View_Log"
            Else
               Str = "Select * From View_Log where (ID like '%" & Value & "%') or (SXDW like '%" & Value & "%') or (SBLX like '%" & Value & "%') or (JQXH like '%" & Value & "%') or (SL like '%" & Value & "%') or (SBBH like '%" & Value & "%') or (XLH like '%" & Value & "%') or (LXRY like '%" & Value & "%') or (LXDH like '%" & Value & "%')" & _
                     " or (SBPH like '%" & Value & "%') or (FZ like '%" & Value & "%') or (SBZT like '%" & Value & "%') or (WXFS like '%" & Value & "%') or (GZXX like '%" & Value & "%') or (SXRK like '%" & Value & "%') or (WXRY like '%" & Value & "%') or (ZT like '%" & Value & "%') or (GHBJ like '%" & Value & "%') or (BZ like '%" & Value & "%') or (QJRY like '%" & Value & "%') or (Account like '%" & Value & "%') or (OperaName like '%" & Value & "%')"
            End If
'            Else '�豸����
'                  Str = "Select * From View_Maintain where SBPH='" & Value & "' order by ID desc"
'            End If
         Else
            If Value = "" Then
'               Str = "Select * From View_Log where (SXRQ between '" & BeginDate & "' and '" & EndDate & "') or (XHRQ between '" & BeginDate & "' and '" & EndDate & "') or (QJRQ between '" & BeginDate & "' and '" & EndDate & "') order by ID desc"
               Str = "Select * From View_Log where (OperaTime between '" & BeginDate & "' and '" & EndDate & "')"
            Else
               Str = "Select * From View_Log where ((ID like '%" & Value & "%') or (SXDW like '%" & Value & "%') or (SBLX like '%" & Value & "%') or (JQXH like '%" & Value & "%') or (SL like '%" & Value & "%') or (SBBH like '%" & Value & "%') or (XLH like '%" & Value & "%') or (LXRY like '%" & Value & "%') or (LXDH like '%" & Value & "%')" & _
                     " or (SBPH like '%" & Value & "%') or (FZ like '%" & Value & "%') or (SBZT like '%" & Value & "%') or (WXFS like '%" & Value & "%') or (GZXX like '%" & Value & "%') or (SXRK like '%" & Value & "%') or (WXRY like '%" & Value & "%') or (ZT like '%" & Value & "%') or (GHBJ like '%" & Value & "%') or (BZ like '%" & Value & "%') or (QJRY like '%" & Value & "%') or (Account like '%" & Value & "%') or (OperaName like '%" & Value & "%'))" & _
                     " and (OperaTime between '" & BeginDate & "' and '" & EndDate & "')"
            End If
         End If

    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetLogRs = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20010
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
Public Function GetLogIDRs(ByVal ID As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Select * From View_Log where ID='" & ID & "'"
   
    
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetLogIDRs = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20011
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
'��ά����Ա��������ȡEmpID
Public Function GetEmpID(ByVal Name As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
       Str = "Select * From Account where Name='" & Name & "'"
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetEmpID = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20012
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'���豸���ͣ���ȡrecordset
Public Function GetDevClsIDRs(ByVal DeviceClass As String) As Recordset
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Select * From DeviceClass where DeviceClass='" & DeviceClass & "'"
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    Set GetDevClsIDRs = ObjDML.ExeSelect(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20013
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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
    
'�����ʺ�
Public Function AddDeviceClass(ByVal DeviceClass As String, ByVal DevAttribute As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Insert Into DeviceClass (DeviceClass,Attribute) Values ('" & DeviceClass & "','" & DevAttribute & "')"
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    AddDeviceClass = ObjDML.ExeInsert(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20014
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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

'�����豸����
Public Function UpdateDeviceClass(ByVal OldDevCls As String, ByVal DeviceClass As String, ByVal DevAttribute As String) As Long
    Dim ObjDML As Object
    Dim Str As String
    Dim ErrNum As Long
    
    ErrorNo = 0
    
    Str = "Update DeviceClass set DeviceClass='" & DeviceClass & "',Attribute='" & DevAttribute & "' Where DeviceClass='" & OldDevCls & "'"
    
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = CreateObject("Com_DML.clsDML")
    On Error GoTo ErrorHandlerKnow
    UpdateDeviceClass = ObjDML.ExeUpdate(Str, DbClass)
    On Error GoTo ErrorHandlerUnknow
    Set ObjDML = Nothing
    Exit Function

ErrorHandlerUnknow:
    'δ֪����
    ErrorNo = 20015
    Err.Clear
ErrorHandlerKnow:
    '��֪����
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