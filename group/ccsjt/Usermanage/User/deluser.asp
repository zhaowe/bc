<%
   
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn%>
<%
'if session("loginid")=nothing then
'Response.Redirect"login.htm"
%>
<%
logid=session("loginid")
objrst.Source ="delete from szairlineuser where logid='"& logid &"'"
objrst.Open 

on error resume next
dim Userid
Userid=Request.QueryString ("Userid")
dim objdml
set objdml=server.CreateObject("Com_UserManage1.clsUserManage1")
ierror=objdml.DelUser (UserID)
if Err.number <>0 then
ierror=Err.number 
Err.Clear 
set objdml=nothing
Response.Redirect "../../Sorry.asp?Errorno="&ierror
end if
set objdml=nothing 
Response.Redirect "userinfo.asp"
%>