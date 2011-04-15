<!--#include file="public.inc"--->
<%
dim groupid
dim groupinfo
dim objdml
dim ierrno
dim groupuser
dim locale
dim i
dim howmanyfields
    groupid=Request.QueryString ("which")
    session("groupid")=groupid
set objdml=server.createobject("Com_UserManage1.clsUserManage1")

on error resume next
set groupinfo=server.createobject("adodb.recordset")
set groupinfo=objdml.GetGroupInfo(groupid,locale)
set  groupuser=objdml.GetGroupUser(groupid,application("UseObject"),locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
howmanyfields=groupuser.fields.count-1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="style.css">
<title>删除组信息</title>
</head>


<body bgcolor="#FFFFFF" style="FONT-SIZE: 10.5pt">
<p><b><font color="blue">删除组信息</font></b> </p>
<table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor="#000000">
  <tr> 
    <td width="100" align="right" bgcolor="#003333"><font color="#FFFFFF">删除的ID:</font></td>
    <td width="489" bgcolor="#FFFFFF"><% =groupinfo("groupid")%></td>
  </tr>
  <tr> 
    <td width="100" align="right" bgcolor="#003333"><font color="#FFFFFF">组名:</font></td>
    <td width="489" bgcolor="#FFFFFF"><% =groupinfo("groupname")%></td>
  </tr>
  <tr> 
    <td width="100" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
    <td width="489" bgcolor="#FFFFFF"><% =groupinfo("description")%></td>
  </tr>
  <%set grouupinfo=nothing%> 
</table>                                   
 
<p>&nbsp;</p>
<p><font color="#0000FF"><b>组中的所有用户 </b></font></p>
<table border=0 width="610" cellPadding=4  cellSpacing=1 bgcolor="#000000">
  <tr bgcolor="#003333"> <%  
for j=2 to 6 
%> <b> 
    <td valign=top border="1" width="298"> <font color="#FFFFFF"><%=groupuser(j).name%> 
      </font></TD>
    <%  
next  
for j=9 to howmanyfields
%> 
    <td valign=top border="1" width="301"> <font color="#FFFFFF"><%=groupuser(j).name%> 
      </font></td>
    </b> <%  
next  
%> </tr>
  <tr bgcolor="#FFFFFF"> <% do while not groupuser.eof %> <%  
for j=2 to 6 
%> 
    <td valign=top border="1" width="298"> <%=groupuser(j)%></td>
    <%  
next  
%> <%  
for j=9 to howmanyfields
%> 
    <td valign=top border="1" width="301"> <%=groupuser(j)%></td>
    <%  
next  
%> </tr>
  <% 

groupuser.movenext
loop
groupuser.close  
set objdml=nothing
%> 
</table>
             
<p>[ <A href="delgroupuser1.asp">删除</a> ] [ <A href="groupinfo.asp">放弃删除</a> ]</p>
</BODY></HTML>
