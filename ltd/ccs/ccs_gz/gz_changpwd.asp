<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�޸�����</title>
</head>

<body>

  
 <p align="center">�� </p>


 

 <p align="center"><font face="����" size="6" color="#6666ff"><b>�޸�����</b></font> </p>


<div align="center">
  <center>
  <form method="post" action="gz_changpwd.asp"  id=form1 name=form1>
  <table border="1" width="46%" bordercolor="#ff99ff">
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">�û�����</font><input  name="loginid">
    </td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">�����룺</font><input type="password" name="password0" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">�����룺</font><input type="password" name="password1" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">�����룺</font><input type="password" name="password2" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><input type="submit" value="�ύ" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <input type="reset" value="ȡ��" name="B2">
      <input type="hidden" value=1 name="tj">  
    </td> 
      
  </tr> 
  <tr>
    <td width="100%" align="middle"><font color=red> ���������Ӣ����ĸ�����ֵ���ϣ�<br> 8&lt;=���볤��&lt;=10�������벻�����������ͬ�� </font></td>
  </tr>  
</table> 
</form> 
<p>
[ <A href="javascript:history.back()">����</A> ]
</p>
  </center> 
</div> 
 
<%
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn
   tj=Request.Form("tj")
   
   if tj=1 then
   
dim loginid
loginid=Request.Form("loginid")
dim password0,password1,password2

password0=Request.Form("Password0") 
password1=Request.Form("Password1") 
password2=Request.Form("Password2")


if lcase(trim(session("loginid")))<>lcase(trim(password1)) then
if password1<>"" then
   if len(cstr(password1))<=10 then
     if len(trim(password1))>=8 then
       if password1=password2 then
       
          
       '����ĳ�8-10��������ĸ���
       
        temppwd=password1
        temppwd=ucase(trim(temppwd))
        havechar=0
        havenum=0
        length=len(password1)
        
        
        
        for i=0 to length
          
          temppwd=left(temppwd,length-i)
          current=right(temppwd,1)
          if current>="0" and  current<="9" then
             havenum=1
          'else
          end if
          if current>="A" and  current<="Z" then
             havechar=1
          '    else
          '       Response.Write "���벻�����г�Ӣ����ĸ������������ַ���"                 
          end if   
          'end if
         next
        
        
        
        
        
        if havenum+havechar=2 then
           
       
       '����ĳ�8-10��������ĸ���
       
        
      
                    
          objRst.Source="select loginid,password from userinfo where loginid='"& trim(LoginID) &"' and password='"& trim(password0) &"' "
          'Response.Write objrst.Source
          'Response.End
          objRst.Open 
          
          if not (objrst.EOF or objrst.BOF) then       
          
          objrst.Close   
          objRst.Source="select loginid,password from userinfo where loginid='"& trim(LoginID) &"' and password='"& trim(password1) &"' "
          objRst.Open 
 
          if objrst.EOF and objrst.BOF then       
        '�ж��¾������Ƿ�һ��
        
        
         objrst.Close
         objrst.Source="update userinfo set password='" & trim(password1)  & "' where loginid='" & trim(LoginID) & "'"
         objrst.Open
         
         Response.Write("<script>window.alert('�޸�����ɹ�!')</script>") '"�޸�����ɹ�!"
         
        
        '�¾��������һ��
         else
            Response.Write "�������ԭ�������벻��һ����"      
         end if 
         
        else
            Response.Write "�������˴���ľ�������û�����"                 
        end if
        
         else
            Response.Write "���������Ӣ����ĸ�����ֵ���ϣ�"                 
         end if
             
       
       else 
         Response.Write "�Բ���������������벻��ȷ������<����>��ť�������룡"
       end if
      else
       Response.Write "�Բ���������������벻������8λ��" 
      end if
     else
       Response.Write "�Բ���������������벻�ܶ���10λ��"
   end if
 else
   Response.Write "�Բ������ٴ��������������룡������������벻��Ϊ�գ�"
end if
else
   Response.Write "�Բ���������û���������ȫһ����"
end if
   
   end if
%>


</body> 
 
</html> 
