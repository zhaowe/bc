<!--#include file="public.inc"--->
<%
dim groupid
dim dbjdml
dim objrs
dim ierrno
    groupid=session("groupid")
    locale=Request.QueryString("locale")
  
set objdml=server.createobject("Com_UserManage1.clsUserManage1")
on  error resume next
 groupname=objdml.GetGroupName(groupid,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改组子表信息</title>
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<b><font color="blue">修改子表信息</font></b> <form name=editgrouplocale action="editgrouplocale.asp" method="post"> 
<input type=hidden name="groupid" value="<%=groupid %>">
 <input type=hidden name="locale" value="<% =locale %>">
<table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">组ID:</font></td>
    <td width="" bgcolor="#FFFFFF"><% =groupid%></td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">版本号</font></td>
    <td width="" bgcolor="#FFFFFF"><%=locale %> </td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">组名</font></td>
    <td width="" bgcolor="#FFFFFF"> 
      <input type=test name="groupname" value="<% =groupname%>" maxlength=50>
    </td>
  </tr>
</table> 
<br>
<input type=submit name=button value="提交"><input type=reset name=button2 value="重置">
 
<input type="button" name="reset_button2" value="返回" onclick="self.history.back()">             
                                                     

 