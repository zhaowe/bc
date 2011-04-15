VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "用户记录"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDbTime 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdGetTime 
      Caption         =   "Get DB Time"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtGuid 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   3930
      Width           =   3615
   End
   Begin VB.CommandButton CmdGuid 
      Caption         =   "Get Guid"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSpTest 
      Caption         =   "SP Test "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "insert"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "delete"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "显示纪录"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "update"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Text            =   "test1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "7"
      Top             =   2400
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
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
   Begin VB.Label Label1 
      Caption         =   "id"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "name"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "DbClass"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBstr As String

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGetTime_Click()
    Dim DBTime As Date
    Dim obj As Object
    
    Set obj = CreateObject("Com_DML.clsDML")
    DBTime = obj.GetDBTime()
    txtDbTime.Text = CStr(DBTime)
End Sub

Private Sub CmdGuid_Click()
    Dim Guid As String
    Dim obj As Object
    
    Set obj = CreateObject("Com_DML.clsDML")
    Guid = obj.GetGuid()
    txtGuid.Text = Guid
    
End Sub


'test SP
Private Sub cmdSpTest_Click()
    Dim objSp As Object
    Dim Param0 As ADODB.Parameter
    Dim Param1 As ADODB.Parameter
    Dim Param2 As ADODB.Parameter
    Dim Param3 As ADODB.Parameter
    Dim Param4 As ADODB.Parameter
    Dim Test1 As Long
    Dim Test2 As ADODB.Parameter
    Dim TestValue As Variant
    Dim Rs As Variant
    
    On Error GoTo ErrorHandler
    Set objSp = CreateObject("Com_DML.clsSP")
    'DbClass = "provider=sqloledb;server=10.101.3.90;database=netbookingtest;uid=sa;" 'test+
    
    Set Param0 = objSp.ParamAdd(, adInteger, adParamReturnValue)
    Set Param1 = objSp.ParamAdd(, adInteger, adParamInput, , 5)
    Set Param2 = objSp.ParamAdd(, adBSTR, adParamInput, 3, "aaa")
    Set Param3 = objSp.ParamAdd(, adInteger, adParamOutput)
    Set Param4 = objSp.ParamAdd(, adBSTR, adParamOutput, 5)
    
    DBstr = Text3.Text
    Set Rs = objSp.exeSp("Sptest", DBstr)
    'if rs is recordset
    Do Until Rs.EOF
        Debug.Print Rs(0)
        Debug.Print Rs(1)
        Rs.MoveNext
    Loop
    
    Debug.Print Param0.Value
    Debug.Print Param3.Value
    Debug.Print Param4.Value
    
    objSp.ParamRefresh
    Test1 = objSp.ParamCount
    Set Test2 = objSp.ParamItem(0)
    TestValue = Test2.Value
    objSp.ParamDel (1)
    Set objSp = Nothing
    
    Debug.Print Test1
    Debug.Print TestValue

    Exit Sub
    
ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description
    
End Sub


Private Sub cmdInsert_Click()
    
    Dim obj As Object
    Dim sStr As String
    Dim iErrNo As Long
    Dim iErrNoVar As Variant
    Dim strSql As String
    Dim str1 As String
    Dim str2 As String
    Dim ino As Integer
    
    str1 = Text1.Text
    str2 = Text2.Text
    ino = CInt(str1)
    
    On Error GoTo ErrorHandler
    Set obj = CreateObject("Com_DML.clsDML")
    
    strSql = "insert into test (id,name) values (" & ino & ",'" & str2 & "')"
    strSql = "Insert into UserHistory (UserID,LoginID,StartDate,EndDate,Status) values ('{A8E66A9C-C4D2-11D4-8654-00805F594010}','d','2000-11-29 01:31:49','2001-05-29','d')"
    
    DBstr = Text3.Text
    iErrNo = obj.ExeInsert(strSql, DBstr)
    'iErrNoVar = obj.Exesql(strSql, DBstr)
    
    'iErrNo = obj.ExeInsert(strSQL, DBstr)
    'If iErrNo <> 0 Then
    '    MsgBox iErrNo
    'End If
    Call cmdDisplay_Click
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description

End Sub

Private Sub cmdDelete_Click()
    Dim obj As Object
    Dim sStr As String
    Dim iErrNo As Integer
    Dim strSql As String
    Dim str1 As String
    Dim str2 As String
    Dim ino As Integer
    
    str1 = Text1.Text
    str2 = Text2.Text
    ino = CInt(str1)
    
    On Error GoTo ErrorHandler
    Set obj = CreateObject("Com_DML.clsDML")
    
    strSql = "delete from test where id=" & ino & ""
    DBstr = Text3.Text
    iErrNo = obj.ExeDelete(strSql, DBstr)
    If iErrNo <> 0 Then
        MsgBox iErrNo
    End If
    Call cmdDisplay_Click
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description
   
End Sub

Private Sub cmdDisplay_Click()
    
    Dim obj As Object
    Dim objRs As ADODB.Recordset
    Dim sStr As String
    Dim iErrNo As Integer
    Dim strSql As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    Set obj = CreateObject("Com_DML.clsDML")
    
    strSql = "Select UserID,Description from View_UserFunction "
    'strSql = "select * from test order by id"
    'strSQL = "declare @Bookid uniqueidentifier select @Bookid=newid() select @Bookid"
    DBstr = Text3.Text
    Set objRs = New ADODB.Recordset
    Set objRs = obj.ExeSQL(strSql, DBstr)
    Set obj = Nothing
    
    'Debug.Print objRs(0)
    'Set objRs = objRs.NextRecordset
    'Debug.Print objRs(0)
    'objRs.MoveLast
    'Debug.Print objRs(0)
    'Debug.Print objRs("@Bookid")
    
    i = objRs.RecordCount
    
    MsgBox (CStr(i))
    Set DataGrid1.DataSource = objRs
    Set objRs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description
    
End Sub

Private Sub cmdUpdate_Click()
    Dim obj As Object
    Dim sStr As String
    Dim iErrNo As Integer
    Dim strSql As String
    Dim str1 As String
    Dim str2 As String
    Dim ino As Integer

    str1 = Text1.Text
    str2 = Text2.Text
    ino = CInt(str1)

    On Error GoTo ErrorHandler
    Set obj = CreateObject("Com_DML.clsDML")

    strSql = "update test set name='" & str2 & "' where id=" & ino & ""
    DBstr = Text3.Text
    iErrNo = obj.ExeUpdate(strSql, DBstr)
    If iErrNo <> 0 Then
        MsgBox iErrNo
    End If
    Call cmdDisplay_Click
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description

End Sub

Private Sub Command1_Click()
    
    Dim obj As Object
    Dim objRs As ADODB.Recordset
    Dim sStr As String
    Dim iErrNo As Integer
    Dim strSql As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    Set obj = CreateObject("Com_UserManage.ClsUserManage")
    
    'strSql = "Select UserID,Description from View_UserFunction "
    'strSql = "select * from test order by id"
    'strSQL = "declare @Bookid uniqueidentifier select @Bookid=newid() select @Bookid"
    'DBstr = Text3.Text
    Set objRs = New ADODB.Recordset
    Set objRs = obj.SearchUserInfo("AMS", "", "")
    Set obj = Nothing
    
    'Debug.Print objRs(0)
    'Set objRs = objRs.NextRecordset
    'Debug.Print objRs(0)
    'objRs.MoveLast
    'Debug.Print objRs(0)
    'Debug.Print objRs("@Bookid")
    
    i = objRs.RecordCount
    
    
    MsgBox (CStr(i))
    Set DataGrid1.DataSource = objRs
    Set objRs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox Err.Number & "," & Err.Source & "," & Err.Description
    

End Sub

Private Sub Form_Load()
'    On Error GoTo AAA
'    Err.Raise 65536, "aaa", "bbb"
'AAA:
'    MsgBox Err.Number & "," & Err.Source & "," & Err.Description
End Sub
