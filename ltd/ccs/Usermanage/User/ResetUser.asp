
<%
on error resume next
dim userid
userid=Request.QueryString ("userid")
dim objdml
set objdml=server.CreateObject ("Com_UserManage.ClsUserManage")
ierror=objdml.RestoreUser (UserID)
if Err.number <> 0 then
ierror=Err.number
Err.Clear 
set objdml=nothing
Response.Redirect "../../Sorry.asp?Errorno="&ierror
end if
set objdml=nothing
Response.Redirect "userinfo.asp"
%>  