<%@ Language=VBScript %>
<%


if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage1.clsUserManage1")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"����ϵͳ���õ�")
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

<P><A href="user/UserInfo.asp">�û�����</A></P>
<P><A href="department/infoin.asp">���Ź���</A></P>
<P><A href="group/groupInfo.asp">�����</A></P>
<P><A href="function/functionInfo.asp">���ܹ���</A></P>
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
