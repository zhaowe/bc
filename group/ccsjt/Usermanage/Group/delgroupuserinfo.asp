<!--#include file="public.inc"--->
<%

dim obj
dim groupid
dim userid
   groupid=session("groupid")
   userid=Request.QueryString("userid")
set obj=server.CreateObject("Com_UserManage1.clsUserManage1")
'on error resume
Response.Write userid
Response.End
set userinfo=server.CreateObject("adodb.recordset")
set  userinfo=obj.GetUserInfo(userid,locale)
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


<body background="images/bg.gif" style="FONT-SIZE: 10.5pt">

<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
  <% for i=0 to howmanyfields %>
	<TR>
		<TD><%userinfo(i).name%></TD>
		<TD><%userinfo(i)%></TD>
	</TR>
	<%next%>
	
</TABLE>

