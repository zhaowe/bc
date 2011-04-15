<!--#include file="public.inc"-->
<%
dim description
dim groupid        
dim groupname
dim objdml
dim ierrno
dim addgple

description=Request.Form("description")
groupname=Request.Form("groupname")
set objdml=server.createobject("com_usermanage.clsusermanage")
on error resume next
'groupid=objdml.AddGroup(description)
'addgple=objdml.AddGroupLocale(groupid,groupname,Application("Locale"))
GroupID=objdml.AddGroupAll(description,groupname,Application("Locale"))
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	set objdml=nothing
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo	
End If
set funinfo=server.CreateObject("adodb.recordset")
set objFunction=server.createobject("com_usermanage.clsFunction")
set funinfo=objfunction.GetAllFunction(Application("Locale"))
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
	set objdml=nothing
End If
set groupfun=server.CreateObject("adodb.recordset")
set groupfun=objdml.GetGroupFunction(groupid,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
	set objdml=nothing
End If
session("groupid")=groupid
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>增加组</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF" style="FONT-SIZE: 10.5pt">
<p><b><font color="blue">增加的组为:</font></b> </p>
<table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor="#000000">
  <tr> 
    <td width="102" align="right" bgcolor="#003333"><font color="#FFFFFF">组号:</font></td>
    <td width= bgcolor="#FFFFFF" bgcolor="#FFFFFF"><% =groupid%></td>
  </tr>
  <tr> 
    <td width="102" align="right" bgcolor="#003333"><font color="#FFFFFF">组名:</font></td>
    <td width= bgcolor="#FFFFFF" bgcolor="#FFFFFF"><% =groupname%></td>
  </tr>
  <tr> 
    <td width="102" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
    <td width= bgcolor="#FFFFFF" bgcolor="#FFFFFF"><%=Request.Form("description")%></td>
  </tr>
</table>
<br>
 <form name=editgroup action="editgroupfun.asp" method="post">  
 <input type=hidden name="groupname" value="<% =groupname%>"> 
 <input type=hidden name="groupid" value="<% =groupid%>">
  <font color=blue>所有的功能:</font> 
  <table border=0 cellPadding=4  cellSpacing=1 bgcolor="#000000" width="610">
    <TR bgcolor="#FFFFFF"><td>  <% 
  session("count")=funinfo.RecordCount
       dim Functionid
          Functionid="functionid"   
		  Funinfo.movefirst
       for i=0 to funinfo.RecordCount-1 
       dim  func
        func=functionid&i
        groupfun.MoveFirst%> 
      
        <input type=checkbox name="<%=Func%>" value="<% =funinfo("functionid")%>" 
	 <%  for j=0 to groupfun.RecordCount-1%><%if funinfo("functionid")=Groupfun("functionid") then%> checked <%end if%>
	<% 
       groupfun.movenext
      next
	%>>
        <%=funinfo("FunctionName")%>
	<% funinfo.movenext
	next
	%>
	</td></TR>
</TABLE>
  <table width="610" border="0" cellspacing="1" cellpadding="4">
    <tr>
      <td width="418"> 
        <input type="submit" name="submit" value="提交">
        <input type="reset" name="button" value="重置">
        <input type="button" name="reset_button2" value="取消" onClick="self.history.back()">
      </td>
      <td width="192">[ <a href=<%="addgroupuser.asp"&"?which="& groupid%>>增加用户 
        </a> ] | [ <a href="groupinfo.asp">返回</a> ]</td>
    </tr>
  </table>
</form>

</body>
</html>
<%
set funinfo=nothing
set groupfun=nothing
set objdml=nothing
set objFunction=nothing
%>