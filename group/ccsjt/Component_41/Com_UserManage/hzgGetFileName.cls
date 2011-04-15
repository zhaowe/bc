VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HzgAll1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'功能：计数器，当调用时，按照所给的页面名称和操作类型对数据库中相应的记录记增量，并返回增加增量后最新的数值。
'参数：Dbclass=(1-生产系统，20-测试系统）
'返回值：当前数值
Public Function UseCounter(ByVal CountName As String, ByVal CountType As String, ByVal CountStep As Integer, ByVal Dbclass As Integer) As String
    
    Dim DbConnectString, DbString, Cs As String
    Dim ObjUser As New Com_DML.clsDML
    Dim Rs As Recordset
    
    DbString = "select * from Use_counter where ltrim(Rtrim(CountName)) =" + "'" + CountName + "'"
    
    Set Rs = ObjUser.ExeSQL(DbString, Dbclass)
    
    If Not (Rs.BOF And Rs.EOF) Then
       
       cc = Rs(2) + CountStep
       Cs = Str(cc)
       DbString = "Update Use_counter set CountName='" + CountName + "',CountType='" + CountType + "',CountNumber=" + Cs
       Set Rs = ObjUser.ExeSQL(DbString, 20)
    
      Else
       
       DbString = "Insert into Use_counter (CountName,CountType,CountNumber) values (" + "'" + CountName + "','" + CountType + "'," + CountStep + ")"
       Set Rs = ObjUser.ExeSQL(DbString, 20)
       cc = CountStep
       
    End If
        
    UseCounter = Str(cc)
    
End Function




'*****************************************************************************
'录入参数文件名必须在前后有引号。“  ”
'例：
'    Dim ObjUser As New HzgDll.HzgGetFileName
'    Dim aa, bb, cc, dd, ee As String
'
'    aa = Text1.Text
'
'    bb = ObjUser.HzgfileName(aa)
'    cc = ObjUser.HzgFileNameSimple(aa)
'    dd = ObjUser.HzgFileNameEnd(aa)
'    ee = ObjUser.HzgFileNameFore(aa)
'
'    Text2.Text = bb
'    Text3.Text = cc
'    Text4.Text = dd
'    Text5.Text = ee
'


'功能:获取带路径的文件名。
'返回值:带路径的文件名
'输入参数：
Public Function HzgfileName(ByVal AllFileName As String) As String
    Dim mAllfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnKnow
    
    '检查输入値是不为空
    mAllfileName = LTrim(RTrim(AllFileName))
    If Len(mAllfileName) = 0 Then
        HzgfileName = "oooooo"
        Exit Function
    End If
    
    
    midstr = Chr(34)
    StartStrNum = InStr(1, mAllfileName, midstr, 1) + 1
    EndStrNum = InStr(StartStrNum, mAllfileName, midstr, 1) - StartStrNum
    HzgfileName = Mid(mAllfileName, StartStrNum, EndStrNum)
    
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    HzgfileName = "oooooo"
    
End Function


'*****************************************************************************
'功能:获取带路径的文件名前半部分
'返回值:带路径的文件名前半部分
'输入参数：
Public Function HzgFileNameFore(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnKnow
    
    '检查输入値是不为空
    mAllfileName = LTrim(RTrim(AllFileName))
    If Len(mAllfileName) = 0 Then
        HzgFileNameFore = "oooooo"
        Exit Function
    End If
    
    midstr = Chr(34)
    StartStrNum = InStr(1, mAllfileName, midstr, 1) + 1
    EndStrNum = InStr(StartStrNum, mAllfileName, midstr, 1) - StartStrNum
    HzgfileName = Mid(mAllfileName, StartStrNum, EndStrNum)
    
    StartStrNum = 1
    EndStrNum = InStr(1, HzgfileName, ".", 1) - 1
    HzgFileNameFore = Mid(HzgfileName, StartStrNum, EndStrNum)
    
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    HzgFileNameFore = "oooooo"
    
    
End Function

'*****************************************************************************
'功能:获取文件名的后半部分
'返回值:文件名的后半部分
'输入参数：
Public Function HzgFileNameEnd(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnKnow
    
    '检查输入値是不为空
    mAllfileName = LTrim(RTrim(AllFileName))
    If Len(mAllfileName) = 0 Then
        HzgFileNameEnd = "oooooo"
        Exit Function
    End If
    
    midstr = Chr(34)
    StartStrNum = InStr(1, mAllfileName, midstr, 1) + 1
    EndStrNum = InStr(StartStrNum, mAllfileName, midstr, 1) - StartStrNum
    HzgfileName = Mid(mAllfileName, StartStrNum, EndStrNum)
    
    StartStrNum = InStr(1, HzgfileName, ".", 1)
    EndStrNum = Len(HzgfileName)
    HzgFileNameEnd = Mid(HzgfileName, StartStrNum, EndStrNum)
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    HzgFileNameEnd = "oooooo"
    
    
End Function

'*****************************************************************************
'功能:获取文件名前半部分
'返回值:文件名前半部分
'输入参数：
Public Function HzgFileNameSimple(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnKnow
    
    '检查输入値是不为空
    mAllfileName = LTrim(RTrim(AllFileName))
    If Len(mAllfileName) = 0 Then
        HzgFileNameSimple = "oooooo"
        Exit Function
    End If
    
    midstr = Chr(34)
    StartStrNum = InStr(1, mAllfileName, midstr, 1) + 1
    EndStrNum = InStr(StartStrNum, mAllfileName, midstr, 1) - StartStrNum
    HzgfileName = Mid(mAllfileName, StartStrNum, EndStrNum)
    
    StartStrNum = 1
    EndStrNum = InStr(1, HzgfileName, ".", 1) - 1
    HzgfileName = Mid(HzgfileName, StartStrNum, EndStrNum)
    
    StartStrNum = InStrRev(HzgfileName, "\", Len(HzgfileName), 1) + 1
    EndStrNum = Len(HzgfileName)
    HzgFileNameSimple = Mid(HzgfileName, StartStrNum, EndStrNum)
    
    Exit Function
    
ErrorHandlerUnKnow:
    '未知错误
    HzgFileNameSimple = "oooooo"
    
End Function

'*****************************************************************************
'功能:将原文本文件加上加密头，并改成ASP文件
'返回值:改成ASP文件后的文件名
'输入参数：   FileSourceName 需为绝对文件路径,SorryFile 需为Sorry.asp 的相对文件路径

Public Function HzgFileAddSec(ByVal FileSourceName As String, ByVal SorryFile As String, ByVal SecDisp As String) As String

   Dim SecLine001, aa, bb, bbb As String
   
   'FileSourceName = "c:\sss\bbb.txt"
   'FileName = "c:\sss\aaa.txt"
   'SorryFile = Text8.Text
   'SecDisp = Text6.Text
   
   
   '判断是否是文本或超文本文件
   aa = Chr(34) + FileSourceName + Chr(34)
   bb = UCase(Trim(HzgFileNameEnd(aa)))
   bbb = bb
   'HzgFileAddSec = bb
   
   If bb <> ".TXT" And bb <> ".HTML" And bb <> ".HTM" Then
   
      HzgFileAddSec = "oooooo"
      
     Else
     
      aa = Chr(34) + FileSourceName + Chr(34)
      bb = HzgFileNameFore(aa)
      
      FileName = bb + ".asp"
      HzgFileAddSec = bb + ".asp"
      
      If Len(Trim(SorryFile)) = 0 Then
         SorryFile = "../sorry/sorry.asp"
      End If
   
      Fno = FreeFile
      Open FileName For Output As #Fno
   
      Fno1 = FreeFile
      Open FileSourceName For Input As #Fno1
   
   
      '===========================================================================
      '写入加密文件头
      '---------------------------------------------------------------------------
      SecLine001 = "<%@ Language=VBScript %>"
      Print #Fno, SecLine001
      SecLine001 = "<%"
      Print #Fno, SecLine001
      
      
      SecLine001 = "if trim(session(" + Chr(34) + "UID" + Chr(34) + "))<>" + Chr(34) + Chr(34) + " then"
      Print #Fno, SecLine001
      SecLine001 = "   Dim objD"
      Print #Fno, SecLine001
      SecLine001 = "   Set objD = server.CreateObject(" + Chr(34) + "Com_UserManage.ClsUserManage" + Chr(34) + ")"
      Print #Fno, SecLine001
      
      SecLine001 = "   VerifyOk=objD.VerifyUserFunction (session(" + Chr(34) + "UID" + Chr(34) + ")," + Chr(34)
      SecLine001 = SecLine001 + SecDisp
      SecLine001 = SecLine001 + Chr(34) + ")"
      Print #Fno, SecLine001
   
      SecLine001 = "   If VerifyOk = False Then"
      Print #Fno, SecLine001
      SecLine001 = "      session(" + Chr(34) + "errorNo" + Chr(34) + ") = " + Chr(34) + "000002" + Chr(34)
      Print #Fno, SecLine001

      SecLine001 = "      Response.Redirect " + Chr(34) + SorryFile + Chr(34)
      Print #Fno, SecLine001

      SecLine001 = "   End If"
      Print #Fno, SecLine001
   
      SecLine001 = "  Else"
      Print #Fno, SecLine001

      SecLine001 = "   session(" + Chr(34) + "errorNo" + Chr(34) + ") = " + Chr(34) + "000001" + Chr(34)
      Print #Fno, SecLine001

      SecLine001 = "   Response.Redirect " + Chr(34) + SorryFile + Chr(34)
      Print #Fno, SecLine001

      SecLine001 = "End If"
      Print #Fno, SecLine001
    
      SecLine001 = "%>"
      Print #Fno, SecLine001
      
      '===========================================================================
      '如果需处理文件是文本文件，则需执行以下步骤
      '---------------------------------------------------------------------------
      If bbb = ".TXT" Then
      
         SecLine001 = "<HTML>"
         Print #Fno, SecLine001
    
         SecLine001 = "<HEAD>"
         Print #Fno, SecLine001
    
         SecLine001 = "<META NAME=" + Chr(34) + "GENERATOR" + Chr(34) + " Content=" + Chr(34) + "Microsoft Visual Studio 6.0" + Chr(34) + ">"
         Print #Fno, SecLine001
    
         SecLine001 = "</HEAD>"
         Print #Fno, SecLine001
    
         SecLine001 = "<BODY>"
         Print #Fno, SecLine001
    
         SecLine001 = "<pre>"
         Print #Fno, SecLine001
    
      End If
      
      '===========================================================================
      '读取需处理文件内容并写入加密文件
      '---------------------------------------------------------------------------
      Do While Not EOF(Fno1)   ' 检查文件尾。
         Line Input #Fno1, SecLine001   ' 读入一行数据。
         Print #Fno, SecLine001  ' 写入加密文件。
      Loop
   
      
      '===========================================================================
      '如果需处理文件是文本文件，则需执行以下步骤
      '---------------------------------------------------------------------------
      If bbb = ".TXT" Then
      
         SecLine001 = "</pre>"
         Print #Fno, SecLine001
    
         SecLine001 = "</BODY>"
         Print #Fno, SecLine001
    
         SecLine001 = "</HTML>"
         Print #Fno, SecLine001
    
      End If
      
      
      Close #Fno1   ' 关闭文件。
      Close #Fno    ' 关闭文件。
    
      Set fs = CreateObject("Scripting.FileSystemObject")
      fs.deleteFile (FileSourceName)

    
    
  End If

End Function

