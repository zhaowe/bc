
<%
'if session("loginid")=nothing then
'Response.Redirect"login.htm"
%>
<%
on error resume next
dim Userid
Userid=Request.QueryString ("Userid")
dim objdml
set objdml=server.CreateObject("Com_UserManage1.clsUserManage1")
ierror=objdml.DelUser (UserID)
if Err.number <>0 then
ierror=Err.number 
Err.Clear 
set objdml=nothing
Response.Redirect "../../Sorry.asp?Errorno="&ierror
end if
set objdml=nothing 
Response.Redirect "userinfo.asp"
%>