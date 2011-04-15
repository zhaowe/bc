VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6810
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   2520
      TabIndex        =   14
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
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
   Begin VB.CommandButton Command4 
      Caption         =   "œ‘ æÀ˘”–"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtPass 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtLogin 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "µ«¬ººÏ≤È"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdIsExist 
      Caption         =   "ºÏ≤ÈLoginID «∑Ò¥Ê‘⁄"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "ª÷∏¥”√ªß"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "‘›Õ£”√ªß"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "…æ≥˝≤‚ ‘"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "ÃÌº”≤‚ ‘"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "LoginID"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAdd_Click()
    Dim FunctionID As String
    Dim GroupID As String
    Dim User As String
    Dim ObjUser As Object
    Dim Ret As Long
    Dim i As Integer
    Dim FuncID(1)
    Dim Group(1)
    Dim UserID(2)

    Set ObjUser = CreateObject("Com_UserManage.ClsUserManage")
    i = 1
    If i = 1 Then
'        UserID(0) = "{BD2827A5-3875-4B1E-A38F-949882F5CAD6}"
'        UserID(1) = "{EF03E629-B8D5-46FB-8366-63E5DF7E1122}"
'        UserID(2) = "{578E556F-B201-4D9A-B72C-952FCE779861}"
'        FunctionID = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
'        Ret = ObjUser.Putuserfunction(UserID, FunctionID)

        FunctionID = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
        Group(0) = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
        Group(1) = "{7EEE4B65-7367-4537-8270-8443E484C2C0}"
        GroupID = "{3EF5FE8E-7315-4278-98BB-EE1709F49FE0}"
        Ret = ObjUser.PutGroupFunction(GroupID, FunctionID)

'        UserID(0) = "{BD2827A5-3875-4B1E-A38F-949882F5CAD6}"
'        UserID(1) = "{EF03E629-B8D5-46FB-8366-63E5DF7E1122}"
'        UserID(2) = "{578E556F-B201-4D9A-B72C-952FCE779861}"
'        User = "{578E556F-B201-4D9A-B72C-952FCE779861}"
'        GroupID = "{1261DA16-E9EB-4CCF-8559-C51699DCE19F}"
'        Ret = ObjUser.PutuserGroup(User, GroupID)
    Else
        User = "{79A7A74B-A72B-45D5-8200-ED3227C500A6}"
        FuncID(0) = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
        FuncID(1) = "{87B9D4C8-2A44-47A6-9CB5-D004865F5BD7}"
        Ret = ObjUser.PutUserFunction(User, FuncID)

        GroupID = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
        FuncID(0) = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
        FuncID(1) = "{87B9D4C8-2A44-47A6-9CB5-D004865F5BD7}"
        Ret = ObjUser.PutGroupFunction(GroupID, FuncID)

        User = "{79A7A74B-A72B-45D5-8200-ED3227C500A6}"
        Group(0) = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
        Group(1) = "{7EEE4B65-7367-4537-8270-8443E484C2C0}"
        Ret = ObjUser.PutUserGroup(User, Group)
    End If
End Sub

Private Sub CmdDel_Click()
    Dim FunctionID As String
    Dim GroupID As String
    Dim UserID As String
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Long

    FunctionID = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
    'GroupID = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
    GroupID = ""
    Ret = ObjUser.DelGroupFunction(GroupID, FunctionID)

    UserID = ""
    Ret = ObjUser.DelUserFunction(UserID, FunctionID)

    GroupID = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
    Ret = ObjUser.DelUserGroup(UserID, GroupID)
End Sub

Private Sub CmdIsExist_Click()
    Dim LoginID As String
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Boolean

    LoginID = TxtLogin.Text
    Ret = ObjUser.IsExistLoginID(LoginID)
    If Ret = True Then
        MsgBox "True"
    Else
        MsgBox "False"
    End If
End Sub

Private Sub CmdLogin_Click()
    Dim LoginID As String
    Dim Pass, mUseObject As String
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Boolean

    LoginID = TxtLogin.Text
    Pass = TxtPass.Text
    mUseObject = "AMS"
    Ret = ObjUser.CheckLoginID(LoginID, Pass, mUseObject)
    If Ret = True Then
        MsgBox "True"
    Else
        MsgBox "False"
    End If
End Sub

Private Sub CmdRestore_Click()
    Dim UserID As String
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Long

    UserID = "{79A7A74B-A72B-45D5-8200-ED3227C500A6}"
    Ret = ObjUser.RestoreUser(UserID)
End Sub

Private Sub CmdStop_Click()
    Dim UserID As String
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Long

    UserID = "{79A7A74B-A72B-45D5-8200-ED3227C500A6}"
    Ret = ObjUser.PauseUser(UserID)
End Sub

Private Sub Command1_Click()
    Dim ObjUser As New Com_UserManage.ClsUserManage
    Dim Ret As Long
    Dim Rs As Recordset
    '≤‚ ‘AddHistory∫ÕEditHistory
    Dim UserID As String
    Dim GroupID As String
    Dim GroupName As String
    Dim Locale As String
    Dim EndDate As Date
    Dim StrResult As String
    
    UserID = "{E0CDEEF7-0D0D-4AC9-92F2-A856F4822704}"
    Locale = "lye"
    Ret = ObjUser.EditPassword(UserID, Locale)
    Exit Sub
    GroupID = "{D670C6FF-845A-47A6-A6D4-45934B734E04}"
    Locale = "en"
    GroupName = "english"
    Ret = ObjUser.AddGroupLocale(GroupID, GroupName, Locale)
    Exit Sub
    
'    UserID = "{AB41921C-BFBC-4799-9D2C-2FCB11EC960E}"
'    StrResult = ObjUser.GetFuncStr(UserID)
'    MsgBox StrResult
'    Exit Sub
'    Dim LoginID As String
'    Dim UserHistory(2)

'    UserID = "{1FC629F2-0571-4960-902A-1DECB591DEC9}"
'    EndDate = Now()
'    LoginID = "testa"
'    UserHistory(0) = UserID
'    UserHistory(1) = LoginID
'    UserHistory(2) = EndDate
'    'Ret = ObjUser.AddUserHistory(UserHistory)
'    Ret = ObjUser.EditUserHistory(UserID, EndDate)

'    Ret = ObjUser.EditUserEndDate(UserID, EndDate)

'    GroupID = "{6B29182A-9D76-4014-A9CE-466A59506F08}"
'    Set Rs = ObjUser.GetLoginInfo("test")
'    UserID = Rs("UserID")
'    MsgBox Rs.RecordCount
End Sub

Private Sub Command2_Click()
    Dim ObjUser As New Com_UserManage.ClsUserManage
    'Dim ObjUser As Object
    Dim Ret As Long
    Dim UserID As String
    Dim LoginID As String
    Dim FunctionID As String
    Dim StrFunc As String
    Dim Str As String
    Dim Locale As String
    Dim UseObject As String
    Dim Rs As Recordset
    Dim UserInfo(8)
    Dim LoginInfo(4)
    Dim LoginOk As Boolean
'       LoginID    (0)
'       Name       (1)
'       Sex        (2)
'       AgentID    (3)
'       CompanyID  (4)
'       ContactInfo(5)
'       UseObject  (6)
'       password   (7)
'       EndDate    (8)
    
    UserInfo(0) = "li'ang"
    UserInfo(1) = "¡∫'"
    UserInfo(2) = "M"
    UserInfo(3) = "{597D8DAB-994D-4C7D-9ACA-D7694CC1666F}"
    UserInfo(4) = "1"
    UserInfo(5) = "M"
    UserInfo(6) = "et"
    UserInfo(7) = "liang"
    UserInfo(8) = "2001-2-30"
    'Set ObjUser = CreateObject("Com_UserManage.ClsUserManage")
    UserID = ObjUser.AddUser(UserInfo)
    Exit Sub
    
    LoginID = "lye"
    Str = "lye"
    UseObject = "et"
    LoginOk = ObjUser.CheckLoginID(LoginID, Str, UseObject)
    MsgBox LoginOk
    Exit Sub
    
    UserID = "{E0CDEEF7-0D0D-4AC9-92F2-A856F4822704}"
    Str = ObjUser.UserIDToAgentID(UserID)
    MsgBox Str
    Exit Sub
    
    UserID = "{CA64B708-36B0-4F0D-BA30-79FE11E72F40}"
    LoginID = "cccc"
    LoginInfo(0) = "aaaa"
    LoginInfo(1) = "M"
    LoginInfo(2) = "{597D8DAB-994D-4C7D-9ACA-D7694CC1666F}"
    LoginInfo(3) = "1"
    LoginInfo(4) = "3252"
    Str = ObjUser.EditLogin(LoginID, LoginInfo)
    MsgBox UserID
    Exit Sub
'    UserID = "{2708239B-AEFA-4F2E-A435-95BE09A9B4E1}"
'    Ret = ObjUser.DelUser(UserID)
'    MsgBox "Success"
    
'    UserID = "{1010D35B-30EF-4B40-8A67-2E108C75C399}"
'    str = "test214"
'    Set Rs = ObjUser.GetUserGroup(UserID)
'    MsgBox Rs.RecordCount

'    FunctionID = "{4C40CA3A-C217-49FE-8669-7FF0C42E90A6}"
'    Locale = "en"
'    str = "Admini"
'    Ret = ObjUser.AddFunctionLocale(FunctionID, str, Locale)
End Sub

Private Sub Command3_Click()
    'Dim Obj As New ClsUserManage
    Dim Obj As Object
    Dim Ret As String
    Dim FunctionID, GroupID As String
    Dim UserID, ComputerID As String
    Dim Str1, Str2 As String
    
    Set Obj = CreateObject("Com_UserManage.ClsUserManage")
    FunctionID = "{CB3B701F-83EA-4988-916C-2A8F4BB95AE1}"
    GroupID = "{A401B3EB-44E6-49C3-9285-AF11CA21A2FD}"
    UserID = "{DA9B021D-2C38-45E2-AF8A-B266BC9E047B}"
    ComputerID = "{441671B7-26E1-45DC-A1D7-6EFECBBEC0D6}"
    Ret = Obj.GetFuncTogether(UserID, ComputerID)
    MsgBox Ret
End Sub

