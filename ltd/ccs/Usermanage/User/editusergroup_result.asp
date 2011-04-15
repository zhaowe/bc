<%@ Language=VBScript %>
<!--#include file="dbclass.asp"-->
<%
dim userid
userid=Request.QueryString ("userid")
dim objdml
set objdml=server.CreateObject("Com_UserManage.ClsUserManage")
dim b
b=cint(Request.QueryString ("b"))
c=0
GroupID="GroupID"
dim ArrayTest1()
for i=0 to b-1
  Grou=GroupID & cstr(i)
  if Request.Form(Grou) <> ""  then
     c=c+1
    end if
  next
  d=cint(c-1)
redim ArrayTest1(d)
d=0
for i=0 to b-1 
  Grou=GroupID & cstr(i)
  if Request.Form(Grou) <> ""  then
      ArrayTest1(d)=Request.Form(Grou)
      d=d+1
    end if
  next
    ierror=objdml.DelUserGroup(Userid)
    ierror=objdml.PutUserGroup(Userid,ArrayTest1)
    if Err.number <>0 then
		ierror=Err.number 
		Err.Clear
		set objdml=nothing
		response.redirect "../../Sorry.asp?Errorno="&ierror
    end if
set objdml=nothing

 Response.Redirect "EditUsergroup.asp?UserId="&UserId
%>