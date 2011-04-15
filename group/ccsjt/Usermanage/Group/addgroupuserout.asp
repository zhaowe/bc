
<!--#include file="public.inc"-->

<%
dim groupid
dim userid
dim objdml
dim rs
dim ierrno
    groupid=session("groupid")
   session.Abandon
    userid=Request.QueryString ("userid")
set objdml=server.CreateObject("com_usermanage1.clsusermanage1")
on error resume next 
 rs=objdml.PutUserGroup(userid,groupid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nohing
Response.Redirect "editgroup.asp"& "?which="& groupid    
Response.end
%>
