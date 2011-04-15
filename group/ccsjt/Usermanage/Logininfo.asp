<%@ Language=VBScript %>
<%


if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage1.clsUserManage1")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"管理系统配置等")
   if VerifyOk=false then
      session("errorNo")="000002"
      Response.Redirect "../sorry/sorry.asp"
   end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P><A href="user/UserInfo.asp">用户管理</A></P>
<P><A href="department/infoin.asp">部门管理</A></P>
<P><A href="group/groupInfo.asp">组管理</A></P>
<P><A href="function/functionInfo.asp">功能管理</A></P>
<br>



<%
'Response.Write Session("FuncStr")
'Response.Write "<br>"
'Response.Write Session("IntraLoginOk")
'Response.Write "<br>"
'Response.Write Session("UID")
'Response.Write "<br>"
'Response.Write Session("AgentID")
'Response.Write "<br>"
%>



</BODY>
</HTML>
