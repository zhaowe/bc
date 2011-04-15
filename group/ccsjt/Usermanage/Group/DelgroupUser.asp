<!--#include file="public.inc"--->
<%
 
dim obj
dim groupid
dim userid
   groupid=session("groupid")
   userid=Request.QueryString("userid")
     
set obj=server.CreateObject("Com_UserManage1.clsUserManage1")
on error resume next

set userinfo=server.CreateObject("adodb.recordset")
set  userinfo=obj.GetUserInfo(userid,locale,UseObject)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
howmanyfields=userinfo.Fields.Count-1
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="style.css">
<title>删除组信息</title>
</head>


<body bgcolor="#FFFFFF" style="FONT-SIZE: 10.5pt">
<FONT color=blue><STRONG>要从组中删除的用户信息</STRONG></FONT> 
<table border=0 cellPadding=4  cellSpacing=1 width="610">
    <% for i=0 to howmanyfields %> 
    <TR>
		<TD><%=userinfo(i).name%>
		<TD><%=userinfo(i)%></TD>
	</TR>
	<%next%>
	
</TABLE>
<DIV align=center></DIV>
 
<br>
<%my_link1="editgroup.asp" &"?which="& groupid%> <%my_link="delgroupout.asp" &"?userid="& userid%> 
[ <A href="<%=my_link%>">删除</A> ][ <A href="<%=my_link1%>">放弃删除</A> ] 
</body></HTML>


