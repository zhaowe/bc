VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmTest 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   1740
   ClientTop       =   900
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   8085
   Begin VB.CommandButton CmdTemp 
      Caption         =   "test"
      Height          =   375
      Left            =   7440
      TabIndex        =   46
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton CmdMeal 
      Caption         =   "餐食记录集"
      Height          =   375
      Left            =   3120
      TabIndex        =   45
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox TxtJcM 
      Height          =   375
      Left            =   3480
      TabIndex        =   42
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtCName 
      Height          =   375
      Left            =   6000
      TabIndex        =   41
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton CmdJcSzmToName 
      Caption         =   "机场三字码转城市三字码"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton CmdJcSzmToCsSzm 
      Caption         =   "机场三字码转城市三字码"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox TxtCsSzm 
      Height          =   375
      Left            =   6000
      TabIndex        =   36
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TxtJcSzm 
      Height          =   375
      Left            =   3480
      TabIndex        =   35
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TxtUseObject 
      Height          =   375
      Left            =   5040
      TabIndex        =   34
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox TxtCabinCode 
      Height          =   405
      Left            =   2640
      TabIndex        =   32
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdCabin 
      Caption         =   "舱位代码转名称"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TxtCabinName 
      Height          =   375
      Left            =   6480
      TabIndex        =   27
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TxtAirlineCode 
      Height          =   375
      Left            =   4920
      TabIndex        =   26
      Top             =   3120
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid GrdAirCompanyRs 
      Height          =   2295
      Left            =   3960
      TabIndex        =   25
      Top             =   4800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdAirRs 
      Caption         =   "获得航空公司记录集"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox TxtAirCode 
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox TxtAirCompany 
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdAirCodeToName 
      Caption         =   "航空公司代码转名称"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox TxtLocale 
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   4200
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid GrdCityRs 
      Height          =   2295
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdCityRs 
      Caption         =   "获得城市记录集"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdGetArea 
      Caption         =   "城市三字码得Area"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox TxtArea 
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TxtSzm 
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox TxtSzmCity 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtCityName 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdCitySzmToName 
      Caption         =   "城市三字码转名称"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TxtCitySzm 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtTel 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdTelToCitySzm 
      Caption         =   "区号转城市三字码"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "城市名称"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "机场三字码"
      Height          =   255
      Left            =   2520
      TabIndex        =   43
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "机场三字码"
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "城市三字码"
      Height          =   255
      Left            =   4920
      TabIndex        =   38
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "用户对象"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "舱位代码"
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "航空公司代码"
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "舱位名称"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "航空公司名称"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "航空公司代码"
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "语言版本"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "城市三字码"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Area值"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "城市名称"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "城市三字码"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "城市三字码"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "电话区号"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDll As New Com_Base.ClsBase
Dim StrIn As String
Dim strLocale As String
Dim strResult As String
Dim strUseObject As String

'航空公司代码转名称
Private Sub CmdAirCodeToName_Click()
    On Error GoTo Err_Handle
        
    StrIn = TxtAirCode.Text
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    strResult = objDll.AirCodeToName(StrIn, strLocale)
    TxtAirCompany.Text = strResult
    Set objDll = Nothing
    Exit Sub
    
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'获得航空公司记录集
Private Sub CmdAirRs_Click()
    Dim Rs As New ADODB.Recordset
    
    On Error GoTo Err_Handle
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    If TxtUseObject.Text = "" Then
        strUseObject = "csn"
    Else
        strUseObject = TxtUseObject.Text
    End If
    Set Rs = objDll.AirCompanyRs(strUseObject, strLocale)
    Set GrdAirCompanyRs.DataSource = Rs
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'获得舱位名称
Private Sub CmdCabin_Click()
    Dim strIn2 As String
    On Error GoTo Err_Handle
        
    StrIn = TxtCabinCode.Text
    strIn2 = TxtAirlineCode.Text
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    strResult = objDll.CabinCodeToName(StrIn, strIn2, strLocale)
    TxtCabinName.Text = strResult
    Set objDll = Nothing
    Exit Sub
    
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'获得城市记录集
Private Sub CmdCityRs_Click()
    Dim Rs As New ADODB.Recordset
    On Error GoTo Err_Handle
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    Set Rs = objDll.CityRs(, strLocale)
    Set GrdCityRs.DataSource = Rs
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'城市三字码转名称
Private Sub CmdCitySzmToName_Click()
    On Error GoTo Err_Handle
    
    StrIn = TxtSzmCity.Text
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    strResult = objDll.SzmToCityName(StrIn, strLocale)
    TxtCityName.Text = strResult
    Set objDll = Nothing
    Exit Sub
    
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'城市三字码得Area
Private Sub CmdGetArea_Click()
    On Error GoTo Err_Handle
    
    StrIn = TxtSzm.Text
    strResult = objDll.RegionType(StrIn)
    TxtArea.Text = strResult
    Set objDll = Nothing
    Exit Sub
    
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'机场三字码转城市三字码
Private Sub CmdJcSzmToCsSzm_Click()
    On Error GoTo Err_Handle
    
    StrIn = TxtJcSzm.Text
    strResult = objDll.JcSzmToCitySzm(StrIn)
    TxtCsSzm.Text = strResult
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'机场三字码转指定语言城市名
Private Sub CmdJcSzmToName_Click()
    On Error GoTo Err_Handle
    
    StrIn = TxtJcM.Text
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    strResult = objDll.JcSzmToCityName(StrIn, strLocale)
    TxtCName.Text = strResult
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'区号转城市三字码
Private Sub CmdTelToCitySzm_Click()
    On Error GoTo Err_Handle
    
    StrIn = TxtTel.Text
    strResult = objDll.PhoneToCitycode(StrIn)
    TxtCitySzm.Text = strResult
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'获得餐食记录集
Private Sub CmdMeal_Click()
    Dim Rs As New ADODB.Recordset
    On Error GoTo Err_Handle
    If TxtLocale.Text = "" Then
        strLocale = "zh"
    Else
        strLocale = TxtLocale.Text
    End If
    Set Rs = objDll.MealRs(strLocale)
    Set GrdCityRs.DataSource = Rs
    Set objDll = Nothing
    Exit Sub
Err_Handle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CmdTemp_Click()
    Dim Obj As New Com_Base.ClsSzdm
    Dim Rs As Recordset
    Dim StrIn As String
    Dim StrOut As String
    Dim Locale As String
    
    StrIn = "CSN"
    Locale = "ZH"
    Set Rs = Obj.SearchSzdmCityInfo("", "", "ZH")
    MsgBox Rs.RecordCount
End Sub

Private Sub Form_Load()
'    Dim strdate As String
'    strdate = CStr(Date + 100)
'    strdate = Format(strdate, "yyyy/mmm/dd")
'    MsgBox strdate

'测试
'    On Error GoTo Err_Handle
'
'    strIn = "733"
'
'    strResult = objDll.PlaneTypeToPicture(strIn)
'    TxtCsSzm.Text = strResult
'    Set objDll = Nothing
'    Exit Sub
'Err_Handle:
'    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


