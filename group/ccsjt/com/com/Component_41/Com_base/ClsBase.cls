VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ErrorNo As Long
Private ErrorSource As String
Private ErrorDescription As String
Private ObjData As Object
Private objRs As Recordset
Private StrSql As String
Const DbClass = 9 'parameter

'���ܣ��绰����ת��Ϊ�������ִ���
'����ֵ���������ִ���
'�������:PhoneNumber�����е绰����
Public Function PhoneToCitycode(ByVal Phonenumber As String) As String
    On Error GoTo Err_Handle
    ErrorNo = 0
        If Phonenumber = "" Then
        ErrorNo = 10059  '�绰����Ϊ��
        ErrorSource = "Com_Base.Clsbase.PhoneToCitycode"
        ErrorDescription = "TelePhone Region code is null"
        GoTo ErrValue
    End If
    StrSql = "select CsSzdm from View_Phonecity where CityPhone='" & Phonenumber & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        ErrorNo = 10060  '�������ִ��벻����
        ErrorSource = "Com_Base.Clsbase.PhoneToCitycode"
        ErrorDescription = "CityPhone is not exist"
        GoTo ErrValue
    Else
        PhoneToCitycode = objRs("CsSzdm")
    End If
    GoTo ErrValue
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10061 'PhoneToCitycode������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.PhoneToCitycode", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��������ִ���ת��Ϊָ�����Եĳ�������
'����ֵ����������
'�������:CitySzm���������ִ���
Public Function SzmToCityName(ByVal CitySzm As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If CitySzm = "" Then
'        ErrorNo = 10062  '�������ִ���Ϊ��
'        ErrorSource = "Com_Base.Clsbase.SzmToCityName"
'        ErrorDescription = "CitySzm is null"
        SzmToCityName = ""
        GoTo ErrValue
    End If
    StrSql = "select CityName from View_CitySzdmToName where CsSzdm='" & CitySzm & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
'        ErrorNo = 10063  '��Ӧ�ĳ������Ʋ�����
'        ErrorSource = "Com_Base.Clsbase.SzmToCityName"
'        ErrorDescription = "City Szm is not exist"
        SzmToCityName = ""
        GoTo ErrValue
    Else
        SzmToCityName = Trim(objRs("CityName"))
    End If
    GoTo ErrValue
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10064 'SzmToCityName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.SzmToCityName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��ɳ������ִ����Ӧ�ö���õ�������𣺹��ʡ����ڡ�����
'����ֵ���������0���ڡ�1���ʡ�2����
'�������:CitySzm���������ִ���   UseObject��Ӧ�ö���(Csn��Et)
Public Function RegionType(ByVal CitySzm As String, Optional ByVal UseObject As String = "csn") As String
    'Dim CheckResult As Long
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If CitySzm = "" Then
        ErrorNo = 10065  '�������ִ���Ϊ��
        ErrorSource = "Com_Base.Clsbase.RegionType"
        ErrorDescription = "CitySzm is null"
        GoTo ErrValue
    End If
    StrSql = "select Area from UseSzdmCity where CsSzdm='" & CitySzm & "' and UseObject='" & UseObject & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        ErrorNo = 10066  '��Ӧ��������𲻴���
        ErrorSource = "Com_Base.Clsbase.RegionType"
        ErrorDescription = "CitySzm is not exist"
        GoTo ErrValue
    Else
        RegionType = objRs("Area")
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10067 'RegionType������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.RegionType", "other error"
    Exit Function
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ�����ָ�����ָ�����Եõ����м�¼��
'����ֵ�����м�¼��
'�������: UseObject��Ӧ�ö���(Csn��Et��)   Locale��ָ������(en��ja��zh��zh-hk)     AppointTyp��ָ�����(All���ɳ�Ʊ��)
Public Function CityRs(Optional ByVal UseObject As String = "csn", _
                        Optional ByVal Locale As String = "zh", _
                        Optional ByVal AppointType As String = "all") As Recordset
    On Error GoTo Err_Handle
    ErrorNo = 0
    Select Case AppointType
    Case "all"
        '����
        StrSql = "select CsSzdm,CityName,PyCity from View_CityRs where Locale='" & Locale & "' " & _
        "and UseObject='" & UseObject & "' order by PyCity"
    Case "1"
        '����Ʊ
        StrSql = "select CsSzdm,CityName,PyCity from View_CityRs where Locale='" & Locale & "' " & _
        "and UseObject='" & UseObject & "' and SendAllow='Y' order by PyCity"
    Case "other"
        '�Ժ�����
    Case Else
        ErrorNo = 10068  'AppointType�����������
        ErrorSource = "Com_Base.Clsbase.CityRs"
        ErrorDescription = "AppointType is Error"
        GoTo ErrValue
    End Select
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    Set CityRs = objRs
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10069 'CityRs������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.CityRs", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��ж�Localeֵ�Ƿ�Ϸ�
'����ֵ��=0��Localeֵ�Ϸ�����������������
'�������:Locale��ָ������(en��ja��zh��zh-hk)
Public Function LocaleCheck(ByVal Locale As String) As Long
    Dim LcaseLocale
    LcaseLocale = LCase(Locale)
    Select Case LcaseLocale
        Case "zh", "zh-hk", "en", "ja"
            ErrorNo = 0
        Case Else
            ErrorNo = 10070
    End Select
    LocaleCheck = ErrorNo
End Function

'���ܣ��ж�Protocolֵ�Ƿ�Ϸ�
'����ֵ��=0��Protocolֵ�Ϸ�����������������
'�������:Protocol��Э������(http��wap)
Public Function ProtocolCheck(ByVal Protocol As String) As Long
    Dim LcaseProtocol
    LcaseProtocol = LCase(Protocol)
    Select Case LcaseProtocol
        Case "http", "wap"
            ErrorNo = 0
        Case Else
            ErrorNo = 10071
    End Select
    ProtocolCheck = ErrorNo
End Function

'���ܣ��ж�UseObjectֵ�Ƿ�Ϸ�
'����ֵ��=0��UseObjectֵ�Ϸ�����������������
'�������:UseObject��Ӧ�ö���(Csn��Et)
Public Function UseObjectCheck(ByVal UseObject As String) As Long
    Dim LcaseUseobject
    LcaseUseobject = LCase(UseObject)
    Select Case LcaseUseobject
        Case "csn", "et"
            ErrorNo = 0
        Case Else
            ErrorNo = 10072
    End Select
    UseObjectCheck = ErrorNo
End Function

'���ܣ��ɺ��չ�˾����õ�ָ�����Եĺ��չ�˾��
'����ֵ�����չ�˾��
'�������: AirCode�����չ�˾����   Locale��ָ������(en��ja��zh��zh-hk)
Public Function AirCodeToName(ByVal AirCode As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If AirCode = "" Then
        ErrorNo = 10073  '���չ�˾�����������Ϊ��
        ErrorSource = "Com_Base.Clsbase.AirCodeToName"
        ErrorDescription = "AirCode is null"
        GoTo ErrValue
    End If
    StrSql = "select CompanyName from AirlineCompanyLocale where CompanyId='" & AirCode & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Or IsNull(objRs("CompanyName")) Then
        ErrorNo = 10074  '���չ�˾�����Ӧ�ĺ��չ�˾��������
        ErrorSource = "Com_Base.Clsbase.AirCodeToName"
        ErrorDescription = "AirCode is not exist"
        GoTo ErrValue
    Else
        AirCodeToName = Trim(objRs("CompanyName"))
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10075  'AirCodeToName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.AirCodeToName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ������û�����ָ�����Եõ����չ�˾��¼��
'����ֵ�����չ�˾��¼��
'�������: UseObject���û�����(can��et��)   Locale��ָ������(en��ja��zh��zh-hk)
Public Function AirCompanyRs(Optional ByVal UseObject As String = "csn", Optional ByVal Locale As String = "zh") As Recordset
    On Error GoTo Err_Handle
    ErrorNo = 0
    StrSql = "select Companyid,CompanyName,Area from View_Aircompany where Locale='" & Locale & "' and UseObject='" & UseObject & "' order by sortorder"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    Set AirCompanyRs = objRs
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10076 'AirCompanyRs������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.CityRs", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��ɲ�λ���롢���չ�˾���롢ָ�����Եõ���λ��
'����ֵ����λ����
'�������: CabinCode����λ����   AirCode�����չ�˾����   Locale��ָ������(en��ja��zh��zh-hk)
Public Function CabinCodeToName(ByVal CabinCode As String, ByVal AirCode As String, Optional ByVal Locale As String = "zh") As String
    'Dim CheckResult As Long
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If AirCode = "" Or CabinCode = "" Then
        ErrorNo = 10091  '��λ����򺽿չ�˾�����������Ϊ��
        ErrorSource = "Com_Base.Clsbase.CabinCodeToName"
        ErrorDescription = "CabinCode or AirCode is null"
        GoTo ErrValue
    End If
    StrSql = "select CabinName from CabinLocale where CabinId='" & CabinCode & "' and CompanyId='" & AirCode & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Or IsNull(objRs("CabinName")) Then
        ErrorNo = 10092  '��Ӧ�Ĳ�λ��������
        ErrorSource = "Com_Base.Clsbase.CabinCodeToName"
        ErrorDescription = "CabinCode is not exist"
        GoTo ErrValue
    Else
        CabinCodeToName = Trim(objRs("CabinName"))
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10093  'CabinCodeToName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.CabinCodeToName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��������ִ���ת��Ϊ�������ִ���
'����ֵ������������
'�������:JcSzm���������ִ���
Public Function JcSzmToCitySzm(ByVal JcSzm As String) As String
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If JcSzm = "" Then
        ErrorNo = 10133  '�������ִ���Ϊ��
        ErrorSource = "Com_Base.Clsbase.JcSzmToCitySzm"
        ErrorDescription = "JcSzm is null"
        GoTo ErrValue
    End If
    StrSql = "select CsSzdm from SzdmAirport where JcSzdm='" & JcSzm & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
'        ErrorNo = 10134  '��Ӧ�ĳ��������벻����
'        ErrorSource = "Com_Base.Clsbase.JcSzmToCitySzm"
'        ErrorDescription = "CitySzm is not exist"
'        GoTo ErrValue
        JcSzmToCitySzm = ""
    Else
        JcSzmToCitySzm = objRs("CsSzdm")
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10135 'JcSzmToCitySzm������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.JcSzmToCitySzm", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ��������ִ���ת��Ϊ��������
'����ֵ����������
'�������:JcSzm���������ִ���
Public Function JcSzmToCityName(ByVal JcSzm As String, Optional ByVal Locale As String = "zh") As String
    Dim CsSzdm As String
    
    ErrorNo = 0
    On Error GoTo Err_HanDleDll
    CsSzdm = JcSzmToCitySzm(JcSzm)
    JcSzmToCityName = Trim(SzmToCityName(CsSzdm, Locale))
    Exit Function
    
'�����Ĵ���
Err_HanDleDll:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'���ܣ�ָ�����Եõ���ʳ��¼��
'����ֵ����ʳ��¼��
'�������: Locale��ָ������(en��ja��zh��zh-hk)
Public Function MealRs(Optional ByVal Locale As String = "zh") As Recordset
    ErrorNo = 0
    StrSql = "select ServeId,ServeName from specialservice where Locale='" & Locale & "' order by ServeId"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    Set MealRs = objRs
    Set ObjData = Nothing
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
End Function

'���ܣ��������ִ���ת��Ϊ��������
'����ֵ����������
'�������:JcSzm���������ִ���
Public Function JcSzmToJcName(ByVal JcSzm As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If JcSzm = "" Then
        ErrorNo = 10317  '�������ִ���Ϊ��
        ErrorSource = "Com_Base.Clsbase.JcSzmToJcName"
        ErrorDescription = "JcSzm is null"
        GoTo ErrValue
    End If
    StrSql = "select AirportName from AirportLocale where JcSzdm='" & JcSzm & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        ErrorNo = 10318  '��Ӧ�Ļ�����������
        ErrorSource = "Com_Base.Clsbase.JcSzmToJcName"
        ErrorDescription = "AirportName is not exist"
        GoTo ErrValue
    Else
        JcSzmToJcName = Trim(objRs("AirportName"))
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    ErrorNo = 10319 'JcSzmToJcName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.JcSzmToJcName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ�ָ�����ԵĻ��Ҵ���ת��Ϊ��������
'����ֵ����������
'�������:CurrencyId�����Ҵ���
Public Function CurrencyIdToName(ByVal CurrencyId As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    StrSql = "select * from CurrencyLocale where CurrencyId='" & CurrencyId & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        CurrencyIdToName = CurrencyId
    Else
        If objRs("CurrencyName") = "" Or IsNull(objRs("CurrencyName")) Then
            CurrencyIdToName = CurrencyId
        Else
            CurrencyIdToName = Trim(objRs("CurrencyName"))
        End If
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10362 'CurrencyIdToName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.CurrencyIdToName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ�ָ�����Եķɻ�����ת��Ϊ�ɻ�����
'����ֵ���ɻ�����
'�������:PlaneType���ɻ�����
Public Function PlaneTypeToName(ByVal PlaneType As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    StrSql = "select * from PlaneTypeLocale where PlaneType='" & PlaneType & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        PlaneTypeToName = PlaneType
    Else
        If objRs("PlaneTypeName") = "" Or IsNull(objRs("PlaneTypeName")) Then
            PlaneTypeToName = PlaneType
        Else
            PlaneTypeToName = Trim(objRs("PlaneTypeName"))
        End If
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10363 'PlaneTypeToName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.PlaneTypeToName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ�����ָ�����Եķɻ����͵õ��ɻ�ͼƬ·��
'����ֵ���ɻ�ͼƬ·��
'�������:PlaneType���ɻ�����
Public Function PlaneTypeToPicture(ByVal PlaneType As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    StrSql = "select * from PlaneTypeLocale where PlaneType='" & PlaneType & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    
    On Error GoTo Err_Handle
    If objRs.EOF Then
        PlaneTypeToPicture = ""
    Else
        If objRs("PlaneTypePicture") = "" Or IsNull(objRs("PlaneTypePicture")) Then
            PlaneTypeToPicture = ""
        Else
            PlaneTypeToPicture = Trim(objRs("PlaneTypePicture"))
        End If
    End If
    GoTo ErrValue
    
'��������������
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10364 'PlaneTypeToName������������
    Err.Raise ErrorNo, "Com_Base.Clsbase.PlaneTypeToName", "other error"
    Exit Function
    
'�𴦲����Ĵ���
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'�Զ������
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'���ܣ����ض�Ӧ���԰汾��ʹ�ö���Ŀ��Ƴ��м�¼��
'����ֵ�����Ƴ��м�¼��
'�������:UseObject:�û�����  Locale:���԰汾
Public Function ControlCityRs(ByVal UseObject As String, ByVal Locale As String) As Recordset
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select * from View_IaCity where ControlStatus='Y' and Locale='" & Locale & "' and " _
                & "UseObject='" & UseObject & "' order by PyCity"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set ControlCityRs = ObjData.ExeSelect(StrSql, DbClass)
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "ControlCityRs() Error"
End Function

'���ܣ�����ĳ��CityID��Ӧʹ�ö���Ŀ��Ƴ���CityID
'����ֵ��ControlCityID
'�������:CityID:����ID  UseObject:�û�����
Public Function ControlCityID(ByVal CityID As String, ByVal UseObject As String) As String
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select ControlCityId from iaCity where CityID='" & CityID & "' and UseObject='" & UseObject & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    On Error GoTo Err_Handle
    If objRs.EOF Then
        ControlCityID = ""
    Else
        ControlCityID = objRs("ControlCityId")
    End If
    Set objRs = Nothing
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not objRs Is Nothing Then
        Set objRs = Nothing
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "ControlCityID() Error"
End Function

'���ܣ����ض�Ӧ���԰汾��ʹ�ö����֪ͨ���м�¼��
'����ֵ��֪ͨ���м�¼��
'�������:UseObject:�û�����  Locale:���԰汾
Public Function InformCityRs(ByVal UseObject As String, ByVal Locale As String) As Recordset
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select * from View_IaCity where InformStatus='Y' and Locale='" & Locale & "' and " _
                & "UseObject='" & UseObject & "' order by PyCity"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set InformCityRs = ObjData.ExeSelect(StrSql, DbClass)
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "InformCityRs() Error"
End Function

'���ܣ����ض�Ӧ���԰汾��ʹ�ö������Ʊ���м�¼��
'����ֵ����Ʊ���м�¼��
'�������:UseObject:�û�����  Locale:���԰汾
Public Function SendCityRs(ByVal UseObject As String, ByVal Locale As String) As Recordset
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select * from View_IaCity where SendStatus='Y' and Locale='" & Locale & "' and " _
                & "UseObject='" & UseObject & "' order by PyCity"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set SendCityRs = ObjData.ExeSelect(StrSql, DbClass)
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "SendCityRs() Error"
End Function

'���ܣ����ض�Ӧ���Ե���Ʒ��������
'����ֵ��GoodsTypeToName
'�������:GoodsType:��Ʒ���  Locale:���԰汾
Public Function GoodsTypeToName(ByVal GoodsType As String, ByVal Locale As String) As String
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select GoodsTypeName from GoodsTypeLocale where GoodsType='" & GoodsType & "' and " _
                & "Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    On Error GoTo Err_Handle
    If objRs.EOF Then
        GoodsTypeToName = ""
    Else
        GoodsTypeToName = objRs("GoodsTypeName")
    End If
    Set objRs = Nothing
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not objRs Is Nothing Then
        Set objRs = Nothing
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "GoodsTypeToName() Error"
End Function

'���ܣ����ض�ӦCityID�ĳ������ִ���
'����ֵ��CsSzdm(�������ִ���)
'�������:CityID:����ID
Public Function CityIdToSzdm(ByVal CityID As String) As String
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select CsSzdm from SzdmCity where CityID='" & CityID & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    On Error GoTo Err_Handle
    If objRs.EOF Then
        CityIdToSzdm = ""
    Else
        CityIdToSzdm = objRs("CsSzdm")
    End If
    Set objRs = Nothing
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not objRs Is Nothing Then
        Set objRs = Nothing
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "CityIdToSzdm() Error"
End Function

'���ܣ����ض�Ӧ���Ե�CityId�ĳ���
'����ֵ��CityName(��������)
'�������:CityID:����ID  Locale:���԰汾
Public Function CityIdToName(ByVal CityID As String, ByVal Locale As String) As String
    ErrorNo = 0
    On Error GoTo Err_Handle
    StrSql = "select CityName from CityLocale where CityID='" & CityID & "' and Locale='" & Locale & "'"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    On Error GoTo Err_Handle
    If objRs.EOF Then
        CityIdToName = ""
    Else
        CityIdToName = objRs("CityName")
    End If
    Set objRs = Nothing
    Set ObjData = Nothing
    Exit Function
'��������������
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'�Զ���͵����������������Ĵ���
Err_HanDleDll:
    If Err.Number <> 0 Then
        ErrorNo = Err.Number
    End If
    If Not objRs Is Nothing Then
        Set objRs = Nothing
    End If
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    On Error GoTo ErrEnd
    If ErrorNo <> 0 Then
        Err.Raise ErrorNo
    End If
    Exit Function
ErrEnd:
    Err.Raise ErrorNo, "Com_Base.ClsBase", "CityIdToName() Error"
End Function