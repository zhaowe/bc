<!--#include file="public.inc"--->
<%
dim objdml
set objA=server.CreateObject ("Com_UserManage.ClsUserManage")
dim ierror
dim groupname
dim groupid
groupid=Request.Form("groupid")
groupname=Request.Form("groupname")
on error resume next
'onlygroup=ObjA.GetGroupInfo	(groupid,locale)
'groupname=onlygroup("groupname")
groupinfo=objA.EditGroupLocale(groupid,groupname)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing
End If
%>
<%
dim str
dim a
dim functionid
  
dim ArrayTest()
a=session("count")
FunctionID="FunctionID"
groupname=Request.Form("groupname")
description=Request.Form("description")
dim func
set objdml=server.createobject("Com_UserManage.ClsUserManage")
c=0
for i=0 to a-1
 Func=FunctionID & cstr(i)
 d=Request(func)
next
for i=0 to a-1
  Func=FunctionID & cstr(i)
  if Request.Form(Func) <> "" then
     c=c+1
    end if
  next 
  c=cint(c-1)
redim ArrayTest(c)
d=0
for i=0 to a-1
  Func=FunctionID & cstr(i)
  if Request.Form(Func) <> ""  then
      ArrayTest(d)=Request.Form(Func)
      d=d+1
    end if
  next
    ierror=objdml.DelGroupFunction(groupid)
    error=objdml.PutGroupFunction(groupid,ArrayTest)
    if Err.number<>0 then
		iErrNo = Err.number
		Err.Clear
		Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
		set objdml=nothing
	End If
    set objdml=nothing
    Response.Redirect "editgroup.asp?"&"which="&groupid
%>