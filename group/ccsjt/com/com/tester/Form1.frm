VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim obj As ClsFunction
    Dim Str As String
    Set obj = New ClsFunction
    
    FunctionName = "testlp"
    Description = "aaaaaa"
    fFunctionID = ""
    FunctionType = "F"
    Locale = "zh"
    
    Str = obj.AddFunctionAll(FunctionName, Description, fFunctionID, , FunctionType, Locale)
    

End Sub


