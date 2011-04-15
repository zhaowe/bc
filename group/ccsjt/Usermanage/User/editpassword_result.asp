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
Response.Write "修改密码成功!"
else 
Response.Write "对不起，您输入的新密码不正确，请点击<返回>按钮重新输入！"
end if
else
Response.Write "对不起，您输入的新密码不能多于10位！"
end if
else
Response.Write "对不起，请再次输入您的新密码！您输入的新密码不能为空！"
end if
%>
<p>[ <a href="EditPassword.asp?UserID=<%=UserID%>">返回</a> ]</p>
</body>
</html>