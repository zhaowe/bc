<!--#include file="public.inc"-->
<%
dim groupid
dim objdml
    groupid=session("groupid")
set objdml=server.createobject("com_usermanage.clsusermanage")
on error resume next
    str=objdml.DelGroup(groupid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
    session("groupid")=nothing
    Response.Redirect"groupinfo.asp"
session("groupid").Abandon
%>

