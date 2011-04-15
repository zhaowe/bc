<!--#include file="public.inc"-->
<%
dim groupid
dim userid
   groupid=session("groupid")
   userid=Request.QueryString("userid")

  
  
dim objdml
set objdml=server.createobject("Com_UserManage1.clsUserManage1")
on error resume next
    str=objdml.DelUserGroup(userid,groupid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
    Response.Redirect"editgroup.asp?"&"which="&groupid
    session("groupid")=nothing
    session("groupid").Abandon
%>

