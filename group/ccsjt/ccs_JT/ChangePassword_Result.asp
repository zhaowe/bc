<html>
<link rel="stylesheet" href="../../style.css">
<body>
<%
on error resume next
dim username
username=session("username")
dim password1,password2

password1=Session("Password1") 
password2=Session("Password2")


if lcase(trim(session("loginid")))<>lcase(trim(password1)) then
if password1<>"" then
   if len(cstr(password2))<=10 then
     if len(trim(password2))>=8 then
       if password1<>password2 then
       
          
       '����ĳ�8-10��������ĸ���
       
        temppwd=password2
        temppwd=ucase(trim(temppwd))
        havechar=0
        havenum=0
        length=len(password2)
        
        
        
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
       
       
      
      
          Set objConn = Server.CreateObject("ADODB.Connection")
          objConn.Open Application("OledbStr") 
      
          Set objRst=server.CreateObject ("ADODB.Recordset")
          objRst.LockType=3
          objRst.CursorType=3
          set objRst.activeConnection=objConn

          objRst.Source="select loginid,password from userinfo where loginid='"& trim(session("LoginID")) &"' and password='"& trim(password1) &"' "
          objRst.Open 
 
          if objrst.EOF and objrst.BOF then       
        '�ж��¾������Ƿ�һ��
        
        
       
         dim objdml
         set objdml=server.CreateObject ("Com_UserManage1.ClsUserManage1")
         ierror=objdml.EditPassword (username,password2)
         if Err.number <> 0 then
	        ierror=Err.number
	        Err.Clear
	        set objdml=nothing 
	        Response.Redirect "../../Sorry.asp?Errorno=��"&ierror
         end if
         set objdml=nothing
         Response.Write "�޸�����ɹ�!"
         
        
        '�¾��������һ��
         else
            Response.Write "���������벻��ȷ��"      
         end if 
         
        
         else
            Response.Write "���������Ӣ����ĸ�����ֵ���ϣ�"                 
         end if
       

       
       
       else 
         Response.Write "�Բ�������������������������ͬ������<����>��ť�������룡"
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
%>
<p>[ <a href="javascript:history.back()">����</a> ]</p>
</body>
</html>