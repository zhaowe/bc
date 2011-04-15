
<!--#include file="public.inc"-->

<%
dim functionid
    functionid=Request.Form("functionid")
dim conflict
dim objdml
dim ierror
set objdml=server.createobject("com_usermanage.clsFunction")
on error resume next
    str=objdml.delfunction(functionid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
    Response.Redirect "functioninfo.asp"
%>
 