
<%
dim obja
set objA=server.CreateObject ("Com_UserManage.ClsFunction")
dim ierror
dim locale
dim functionid
functionid=Request.Form("functionid")
functionname=Request.Form("functionname")
locale=Request.Form("locale")
on error resume next
functionlocale=obja.EditFunctionLocale( functionid,functionname,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	set objdml=nothing
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
End If
Response.Redirect"fsontablerun.asp?"&"functionid="&functionid
%> 