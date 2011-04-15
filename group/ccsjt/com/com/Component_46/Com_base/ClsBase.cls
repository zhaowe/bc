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

'功能：电话区号转换为城市三字代码
'返回值：城市三字代码
'传入参数:PhoneNumber：城市电话区号
Public Function PhoneToCitycode(ByVal Phonenumber As String) As String
    On Error GoTo Err_Handle
    ErrorNo = 0
        If Phonenumber = "" Then
        ErrorNo = 10059  '电话区号为空
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
        ErrorNo = 10060  '城市三字代码不存在
        ErrorSource = "Com_Base.Clsbase.PhoneToCitycode"
        ErrorDescription = "CityPhone is not exist"
        GoTo ErrValue
    Else
        PhoneToCitycode = objRs("CsSzdm")
    End If
    GoTo ErrValue
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10061 'PhoneToCitycode方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.PhoneToCitycode", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：城市三字代码转换为指定语言的城市名称
'返回值：城市名称
'传入参数:CitySzm：城市三字代码
Public Function SzmToCityName(ByVal CitySzm As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If CitySzm = "" Then
'        ErrorNo = 10062  '城市三字代码为空
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
'        ErrorNo = 10063  '对应的城市名称不存在
'        ErrorSource = "Com_Base.Clsbase.SzmToCityName"
'        ErrorDescription = "City Szm is not exist"
        SzmToCityName = ""
        GoTo ErrValue
    Else
        SzmToCityName = Trim(objRs("CityName"))
    End If
    GoTo ErrValue
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10064 'SzmToCityName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.SzmToCityName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：由城市三字代码和应用对象得到区域类别：国际、国内、地区
'返回值：区域类别：0国内、1国际、2地区
'传入参数:CitySzm：城市三字代码   UseObject：应用对象(Csn、Et)
Public Function RegionType(ByVal CitySzm As String, Optional ByVal UseObject As String = "csn") As String
    'Dim CheckResult As Long
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If CitySzm = "" Then
        ErrorNo = 10065  '城市三字代码为空
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
        ErrorNo = 10066  '对应的区域类别不存在
        ErrorSource = "Com_Base.Clsbase.RegionType"
        ErrorDescription = "CitySzm is not exist"
        GoTo ErrValue
    Else
        RegionType = objRs("Area")
    End If
    GoTo ErrValue
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10067 'RegionType方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.RegionType", "other error"
    Exit Function
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：根据指定类别、指定语言得到城市记录集
'返回值：城市记录集
'传入参数: UseObject：应用对象(Csn、Et等)   Locale：指定语言(en、ja、zh、zh-hk)     AppointTyp：指定类别(All、可出票等)
Public Function CityRs(Optional ByVal UseObject As String = "csn", _
                        Optional ByVal Locale As String = "zh", _
                        Optional ByVal AppointType As String = "all") As Recordset
    On Error GoTo Err_Handle
    ErrorNo = 0
    Select Case AppointType
    Case "all"
        '所有
        StrSql = "select CsSzdm,CityName,PyCity from View_CityRs where Locale='" & Locale & "' " & _
        "and UseObject='" & UseObject & "' order by PyCity"
    Case "1"
        '可送票
        StrSql = "select CsSzdm,CityName,PyCity from View_CityRs where Locale='" & Locale & "' " & _
        "and UseObject='" & UseObject & "' and SendAllow='Y' order by PyCity"
    Case "other"
        '以后添加
    Case Else
        ErrorNo = 10068  'AppointType输入参数错误
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
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10069 'CityRs方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.CityRs", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：判断Locale值是否合法
'返回值：=0，Locale值合法；否则产生错误代码
'传入参数:Locale：指定语言(en、ja、zh、zh-hk)
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

'功能：判断Protocol值是否合法
'返回值：=0，Protocol值合法；否则产生错误代码
'传入参数:Protocol：协议类型(http、wap)
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

'功能：判断UseObject值是否合法
'返回值：=0，UseObject值合法；否则产生错误代码
'传入参数:UseObject：应用对象(Csn、Et)
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

'功能：由航空公司代码得到指定语言的航空公司名
'返回值：航空公司名
'传入参数: AirCode：航空公司代码   Locale：指定语言(en、ja、zh、zh-hk)
Public Function AirCodeToName(ByVal AirCode As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If AirCode = "" Then
        ErrorNo = 10073  '航空公司代码输入参数为空
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
        ErrorNo = 10074  '航空公司代码对应的航空公司名不存在
        ErrorSource = "Com_Base.Clsbase.AirCodeToName"
        ErrorDescription = "AirCode is not exist"
        GoTo ErrValue
    Else
        AirCodeToName = Trim(objRs("CompanyName"))
    End If
    GoTo ErrValue
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10075  'AirCodeToName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.AirCodeToName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：根据用户对象、指定语言得到航空公司记录集
'返回值：航空公司记录集
'传入参数: UseObject：用户对象(can、et等)   Locale：指定语言(en、ja、zh、zh-hk)
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
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10076 'AirCompanyRs方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.CityRs", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：由舱位代码、航空公司代码、指定语言得到舱位名
'返回值：舱位名称
'传入参数: CabinCode：舱位代码   AirCode：航空公司代码   Locale：指定语言(en、ja、zh、zh-hk)
Public Function CabinCodeToName(ByVal CabinCode As String, ByVal AirCode As String, Optional ByVal Locale As String = "zh") As String
    'Dim CheckResult As Long
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If AirCode = "" Or CabinCode = "" Then
        ErrorNo = 10091  '舱位代码或航空公司代码输入参数为空
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
        ErrorNo = 10092  '对应的舱位名不存在
        ErrorSource = "Com_Base.Clsbase.CabinCodeToName"
        ErrorDescription = "CabinCode is not exist"
        GoTo ErrValue
    Else
        CabinCodeToName = Trim(objRs("CabinName"))
    End If
    GoTo ErrValue
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10093  'CabinCodeToName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.CabinCodeToName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：机场三字代码转换为城市三字代码
'返回值：城市三字码
'传入参数:JcSzm：机场三字代码
Public Function JcSzmToCitySzm(ByVal JcSzm As String) As String
    
    On Error GoTo Err_Handle
    ErrorNo = 0
    If JcSzm = "" Then
        ErrorNo = 10133  '机场三字代码为空
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
'        ErrorNo = 10134  '对应的城市三字码不存在
'        ErrorSource = "Com_Base.Clsbase.JcSzmToCitySzm"
'        ErrorDescription = "CitySzm is not exist"
'        GoTo ErrValue
        JcSzmToCitySzm = ""
    Else
        JcSzmToCitySzm = objRs("CsSzdm")
    End If
    GoTo ErrValue
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10135 'JcSzmToCitySzm方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.JcSzmToCitySzm", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：机场三字代码转换为城市名称
'返回值：城市名称
'传入参数:JcSzm：机场三字代码
Public Function JcSzmToCityName(ByVal JcSzm As String, Optional ByVal Locale As String = "zh") As String
    Dim CsSzdm As String
    
    ErrorNo = 0
    On Error GoTo Err_HanDleDll
    CsSzdm = JcSzmToCitySzm(JcSzm)
    JcSzmToCityName = Trim(SzmToCityName(CsSzdm, Locale))
    Exit Function
    
'产生的错误
Err_HanDleDll:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'功能：指定语言得到餐食记录集
'返回值：餐食记录集
'传入参数: Locale：指定语言(en、ja、zh、zh-hk)
Public Function MealRs(Optional ByVal Locale As String = "zh") As Recordset
    ErrorNo = 0
    StrSql = "select ServeId,ServeName from specialservice where Locale='" & Locale & "' order by ServeId"
    Set ObjData = CreateObject("Com_DML.clsDMl")
    On Error GoTo Err_HanDleDll
    Set objRs = ObjData.ExeSelect(StrSql, DbClass)
    Set MealRs = objRs
    Set ObjData = Nothing
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
End Function

'功能：机场三字代码转换为机场名称
'返回值：机场名称
'传入参数:JcSzm：机场三字代码
Public Function JcSzmToJcName(ByVal JcSzm As String, Optional ByVal Locale As String = "zh") As String
    On Error GoTo Err_Handle
    ErrorNo = 0
    If JcSzm = "" Then
        ErrorNo = 10317  '机场三字代码为空
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
        ErrorNo = 10318  '对应的机场名不存在
        ErrorSource = "Com_Base.Clsbase.JcSzmToJcName"
        ErrorDescription = "AirportName is not exist"
        GoTo ErrValue
    Else
        JcSzmToJcName = Trim(objRs("AirportName"))
    End If
    GoTo ErrValue
    
'本方法其它错误
Err_Handle:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    ErrorNo = 10319 'JcSzmToJcName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.JcSzmToJcName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：指定语言的货币代码转换为货币名称
'返回值：货币名称
'传入参数:CurrencyId：货币代码
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
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10362 'CurrencyIdToName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.CurrencyIdToName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：指定语言的飞机类型转换为飞机名称
'返回值：飞机名称
'传入参数:PlaneType：飞机类型
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
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10363 'PlaneTypeToName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.PlaneTypeToName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：根据指定语言的飞机类型得到飞机图片路径
'返回值：飞机图片路径
'传入参数:PlaneType：飞机类型
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
    
'本方法其它错误
Err_Handle:
    Set ObjData = Nothing
    ErrorNo = 10364 'PlaneTypeToName方法其它错误
    Err.Raise ErrorNo, "Com_Base.Clsbase.PlaneTypeToName", "other error"
    Exit Function
    
'别处产生的错误
Err_HanDleDll:
    If Not ObjData Is Nothing Then
        Set ObjData = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
'自定义错误
ErrValue:
    Set ObjData = Nothing
    If ErrorNo <> 0 Then
        On Error GoTo Err_HanDleDll
        Err.Raise ErrorNo, ErrorSource, ErrorDescription
    End If
End Function

'功能：返回对应语言版本和使用对象的控制城市记录集
'返回值：控制城市记录集
'传入参数:UseObject:用户对象  Locale:语言版本
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回某个CityID对应使用对象的控制城市CityID
'返回值：ControlCityID
'传入参数:CityID:城市ID  UseObject:用户对象
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回对应语言版本和使用对象的通知城市记录集
'返回值：通知城市记录集
'传入参数:UseObject:用户对象  Locale:语言版本
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回对应语言版本和使用对象的送票城市记录集
'返回值：送票城市记录集
'传入参数:UseObject:用户对象  Locale:语言版本
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回对应语言的商品类别的名称
'返回值：GoodsTypeToName
'传入参数:GoodsType:商品类别  Locale:语言版本
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回对应CityID的城市三字代码
'返回值：CsSzdm(城市三字代码)
'传入参数:CityID:城市ID
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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

'功能：返回对应语言的CityId的城市
'返回值：CityName(城市名称)
'传入参数:CityID:城市ID  Locale:语言版本
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
'本方法其它错误
Err_Handle:
    ErrorNo = 10618
    Err.Clear
'自定义和调用其它方法产生的错误
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
