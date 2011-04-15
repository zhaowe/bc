<!--#include file="public.inc"-->
<%
dim groupid
dim objdml
dim functionid()
dim i
dim ierrno
dim putgroupfun
groupid=session("groupid")
functionid(0)=Request.form("check1")
functionid(1)=Request.form("check2")

set objdml=server.CreateObject("com_usermanage.clsusermanage")
on error resume next
   putgroupfun=objdml.PutGroupFunction (groupid,functionid())
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If 
set objdml=nothing
session("groupid")=null
Response.Redirect "addgpout.asp"
%>
