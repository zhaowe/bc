VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hzgall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'���ܣ���������������ʱ������������ҳ�����ƺͲ������Ͷ����ݿ�����Ӧ�ļ�¼���������������������������µ���ֵ��
'������Dbclass=num(10.254.0.num 91-����ϵͳ��
'����ֵ����ǰ��ֵ
Public Function UseCounter(ByVal CountName As String, ByVal CountType As String, ByVal CountStep As Integer, ByVal DbClass As Integer) As String
    
    Dim DbConnectString, DbString, Cs As String
    Dim ObjUser As New hzgdll.hzgbase
    Dim Rs As Recordset
    
    DbString = "select * from Use_counter where ltrim(Rtrim(CountName)) ='" + CountName + "'"
    
    Set Rs = ObjUser.ExeSQL(DbString, DbClass)
    
    If Not (Rs.BOF And Rs.EOF) Then
       
       cc = Rs(2) + CountStep
       Cs = Str(cc)
       DbString = "Update Use_counter set CountName='" + CountName + "',CountType='" + CountType + "',CountNumber=" + Cs
       Set Rs = ObjUser.ExeSQL(DbString, DbClass)
    
      Else
       
       DbString = "Insert into Use_counter (CountName,CountType,CountNumber) values (" + "'" + CountName + "','" + CountType + "'," + Str(CountStep) + ")"
       Set Rs = ObjUser.ExeSQL(DbString, DbClass)
       cc = CountStep
       
    End If
        
    UseCounter = Str(cc)
    
End Function




'*****************************************************************************
'¼������ļ���������ǰ�������š���  ��
'����
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


'����:��ȡ��·�����ļ�����
'����ֵ:��·�����ļ���
'���������
Public Function HzgfileName(ByVal AllFileName As String) As String
    Dim mAllfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnknow
    
    '������낎�ǲ�Ϊ��
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
    
ErrorHandlerUnknow:
    'δ֪����
    HzgfileName = "oooooo"
    
End Function


'*****************************************************************************
'����:��ȡ��·�����ļ���ǰ�벿��
'����ֵ:��·�����ļ���ǰ�벿��
'���������
Public Function HzgFileNameFore(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnknow
    
    '������낎�ǲ�Ϊ��
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
    
ErrorHandlerUnknow:
    'δ֪����
    HzgFileNameFore = "oooooo"
    
    
End Function

'*****************************************************************************
'����:��ȡ�ļ����ĺ�벿��
'����ֵ:�ļ����ĺ�벿��
'���������
Public Function HzgFileNameEnd(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnknow
    
    '������낎�ǲ�Ϊ��
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
    
ErrorHandlerUnknow:
    'δ֪����
    HzgFileNameEnd = "oooooo"
    
    
End Function

'*****************************************************************************
'����:��ȡ�ļ���ǰ�벿��
'����ֵ:�ļ���ǰ�벿��
'���������
Public Function HzgFileNameSimple(ByVal AllFileName As String) As String
    Dim mAllfileName, HzgfileName, midstr As String
    Dim i, StartStrNum, EndStrNum As Integer
    
    On Error GoTo ErrorHandlerUnknow
    
    '������낎�ǲ�Ϊ��
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
    
ErrorHandlerUnknow:
    'δ֪����
    HzgFileNameSimple = "oooooo"
    
End Function

'*****************************************************************************
'����:��ԭ�ı��ļ����ϼ���ͷ�����ĳ�ASP�ļ�
'����ֵ:�ĳ�ASP�ļ�����ļ���
'���������   FileSourceName ��Ϊ�����ļ�·��,SorryFile ��ΪSorry.asp ������ļ�·��

Public Function HzgFileAddSec(ByVal FileSourceName As String, ByVal SorryFile As String, ByVal SecDisp As String) As String

   Dim SecLine001, aa, bb, bbb As String
   
   'FileSourceName = "c:\sss\bbb.txt"
   'FileName = "c:\sss\aaa.txt"
   'SorryFile = Text8.Text
   'SecDisp = Text6.Text
   
   
   '�ж��Ƿ����ı����ı��ļ�
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
      'д������ļ�ͷ
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
      '����账���ļ����ı��ļ�������ִ�����²���
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
      '��ȡ�账���ļ����ݲ�д������ļ�
      '---------------------------------------------------------------------------
      Do While Not EOF(Fno1)   ' ����ļ�β��
         Line Input #Fno1, SecLine001   ' ����һ�����ݡ�
         Print #Fno, SecLine001  ' д������ļ���
      Loop
   
      
      '===========================================================================
      '����账���ļ����ı��ļ�������ִ�����²���
      '---------------------------------------------------------------------------
      If bbb = ".TXT" Then
      
         SecLine001 = "</pre>"
         Print #Fno, SecLine001
    
         SecLine001 = "</BODY>"
         Print #Fno, SecLine001
    
         SecLine001 = "</HTML>"
         Print #Fno, SecLine001
    
      End If
      
      
      Close #Fno1   ' �ر��ļ���
      Close #Fno    ' �ر��ļ���
    
      Set fs = CreateObject("Scripting.FileSystemObject")
      fs.deleteFile (FileSourceName)

    
    
  End If

End Function


