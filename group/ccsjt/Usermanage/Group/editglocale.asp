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
<title>�޸����ӱ���Ϣ</title>
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<b><font color="blue">�޸��ӱ���Ϣ</font></b> <form name=editgrouplocale action="editgrouplocale.asp" method="post"> 
<input type=hidden name="groupid" value="<%=groupid %>">
 <input type=hidden name="locale" value="<% =locale %>">
<table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">��ID:</font></td>
    <td width="" bgcolor="#FFFFFF"><% =groupid%></td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">�汾��</font></td>
    <td width="" bgcolor="#FFFFFF"><%=locale %> </td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">����</font></td>
    <td width="" bgcolor="#FFFFFF"> 
      <input type=test name="groupname" value="<% =groupname%>" maxlength=50>
    </td>
  </tr>
</table> 
<br>
<input type=submit name=button value="�ύ"><input type=reset name=button2 value="����">
 
<input type="button" name="reset_button2" value="����" onclick="self.history.back()">             
                                                     

 