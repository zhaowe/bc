<!--#include file="public.inc"-->
<%
dim description
dim conflict
dim functionname
dim objrs
dim  ierrno
dim objdml
dim ffunctionname
on error resume next
    functionname=Request.Form("functionname")
    description=Request.form( "description")
    conflict=Request.form("conflict")
    ffunctionid=Request.Form("ffunctionid")
    functiontype1=Request.Form("functiontype1")
    functiontype2=Request.Form("functiontype2")
    functiontype3=Request.Form("functiontype3")
    functiontype=functiontype1&functiontype2&functiontype3
set Objdml=server.createobject("Com_UserManage.ClsFunction")
str=Objdml.AddFunctionAll(functionname,description,ffunctionid,,functiontype,locale)
if Err.number<>0 then 
	iErrno = Err.number
	Err.Clear
  	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
End If
set ordernum=server.CreateObject ("ADODB.Recordset")
set ordernum=objdml.GetFunctionInfo(str,locale)
if Err.number<>0 then 
	iErrno = Err.number
	Err.Clear
  	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
End If
 set objDML=nothing
 %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>增加功能信息</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<font color="blue"><b>增加的功能信息为:</b></font> <br>
<table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
  <tr bgcolor="#FFFFFF"> 
    <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">功能名:</font></td>
    <td width="369"><% =Request.Form("functionname")%></td>
  </tr>
   <tr bgcolor="#FFFFFF"> 
    <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">序列号:</font></td>
    <td width="369"><% =ordernum("ordernum")%></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
    <td width="369"><% =Request.Form ("description")%></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">父功能名:</font></td>
    <td width="369"><%=ordernum("fFunctionName")%></td>
  </tr>
 <tr bgcolor="#FFFFFF"> 
    <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">功能类型:</font></td>
    <td width="369">
      <%select case functiontype
       case "M"
       Response.Write "菜单"
       case "F"
       Response.Write "功能"  
       case "P"
       Response.Write "页面" 
       case "MF" 
       Response.Write "菜单,功能"
       case "MP"
       Response.Write "菜单,页面"
       case "FP"
       Response.Write "功能,页面"
       case "MFP" 
       Response.Write "功能,菜单,页面"
       End Select
      %>
    </td>
  </tr> 
</table>
<br>
<%set objrs=nothing %><centr><a href="functioninfo.asp">返回</centr> 
</body>
</html>
                                
     
           
