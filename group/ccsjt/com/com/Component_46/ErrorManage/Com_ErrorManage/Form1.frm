VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   795
   ClientTop       =   1350
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9060
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim com_Err As New Com_ErrorManage.ClsErrorManage
Dim inum As Long
Dim rst As New ADODB.Recordset
Dim num As Integer

Private Sub Command1_Click()
    Dim ErrNoBack As Long
    inum = com_Err.ErrorQuery(, Equal, , "y")
    Set rst = com_Err.RsGet
    num = rst.RecordCount
    MsgBox num
    'MsgBox ErrNoBack
End Sub

Private Sub Command2_Click()
    'inum = com_Err.LocaleTypeDeal(10002, zh, wap, "df22", "dfdfh", Delete)
    inum = com_Err.ErrorDeal(101, , "dsf", "sdf", Java, "ss", "dsf", , Inner, Insert)
    MsgBox inum
End Sub

Private Sub Form_Load()
    Set com_Err = New Com_ErrorManage.ClsErrorManage
    Set rst = com_Err.ErrorQuery()
    num = rst.RecordCount
    MsgBox num
    Set DataGrid1.DataSource = rst
End Sub
