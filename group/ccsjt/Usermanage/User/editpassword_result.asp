<html>
<link rel="stylesheet" href="../../style.css">
<body>
<%
on error resume next
dim userid
userid=Request.QueryString("UserID")
dim password1,password2

password1=Request.Form ("password1")
password2=Request.Form ("password2")
if password1<>"" then
if len(cstr(password1))<=10 then
if password1=password2 then
dim objdml
set objdml=server.CreateObject ("Com_UserManage1.clsUserManage1")
ierror=objdml.EditPassword (UserId,password1)
if Err.number <> 0 then
	ierror=Err.number
	Err.Clear
	set objdml=nothing 
	Response.Redirect "../../Sorry.asp?Errorno="&ierror
end if
set objdml=nothing
Response.Write "�޸�����ɹ�!"
else 
Response.Write "�Բ���������������벻��ȷ������<����>��ť�������룡"
end if
else
Response.Write "�Բ���������������벻�ܶ���10λ��"
end if
else
Response.Write "�Բ������ٴ��������������룡������������벻��Ϊ�գ�"
end if
%>
<p>[ <a href="EditPassword.asp?UserID=<%=UserID%>">����</a> ]</p>
</body>
</html>