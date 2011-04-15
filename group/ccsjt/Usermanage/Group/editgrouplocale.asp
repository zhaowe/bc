
<%
dim obja
set objA=server.CreateObject ("Com_UserManage1.clsUserManage1")
dim ierror
dim locale
dim groupid
groupid=Request.Form("groupid")
groupname=Request.Form("groupname")
locale=Request.Form("locale")
on error resume next
grouplocale=obja.EditGroupLocale(groupid,groupname,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set obja=nothing
Response.Redirect"gsontablerun.asp?"&"which="&groupid
%> 