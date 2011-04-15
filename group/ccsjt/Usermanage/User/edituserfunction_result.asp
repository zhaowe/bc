<%@ Language=VBScript %>
<!--#include file="dbclass.asp"-->
    <%'***************************************************
dim userid
userid=Request.QueryString ("userid")
dim objdml
set objdml=server.CreateObject("Com_UserManage1.ClsUserManage1")
dim a
a=cint(Request.QueryString ("a"))
c=0
FunctionID="FunctionID"
dim ArrayTest()
for i=0 to a-1
  Func=FunctionID & cstr(i)
  if Request.Form(Func) <> ""  then     
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
    ierror=objdml.DelUserFunction(Userid)
    ierror=objdml.PutUserFunction(Userid,ArrayTest)
    if Err.number <>0 then
		ierror=Err.number
		Err.Clear
		set objdml=nothing
		response.redirect "../../Sorry.asp?Errorno="&ierror
    end if
    set objdml=nothing
%>       
<%Response.Redirect "EditUserfunction.asp?UserId="&UserId
%>

