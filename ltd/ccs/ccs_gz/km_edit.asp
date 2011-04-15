<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"ccs_gsgly")
   if VerifyOk=false then
      session("errorNo")="000002"
      Response.Redirect "../sorry/sorry.asp"
   end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
%> 
<html>
<head>
<title>预算管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
.px10 {  font-size: 10px; line-height: 150%}
.px12 {  font-size: 12px; line-height: 150%}
.px14 {  font-size: 14px; line-height: 150%}
.px16 {  font-size: 16px; line-height: 150%}
.px18 {  font-size: 18px; line-height: 150%}
.px24 {  font-size: 24px; line-height: 150%}
.px36 {  font-size: 36px; line-height: 150%}
.px48 {  font-size: 48px; line-height: 150%}
.px72 {  font-size: 72px; line-height: 150%}
body {  font-size: 12px; line-height: 150%}
p {  font-size: 12px; line-height: 150%}
td {  font-size: 9px; line-height: 150%}
input {  font-size: 12px; line-height: 150%}
select {  font-size: 12px; line-height: 150%}
.content4{FONT-SIZE:10PT; LINE-HEIGHT:9PT;}
.contentindex{font-family: "宋体";FONT-SIZE:9pt; LINE-HEIGHT:11pt;}
.enter {COLOR: #FFAF02; FONT-FAMILY: "宋体", "Arial", "Times New Roman"; FONT-SIZE: 11pt; TEXT-DECORATION: none ;font-weight: bold}
.head1{FONT-SIZE:11pt; LINE-HEIGHT:18pt; font-weight: bold; }
.head2{FONT-SIZE:10pt; LINE-HEIGHT:14pt; font-weight: bold; }
.contentsmall{FONT-SIZE:9pt; LINE-HEIGHT:12pt;}
.nav{FONT-SIZE:9pt; LINE-HEIGHT:10pt; color: #999999}
.content{FONT-SIZE:10pt; LINE-HEIGHT:14pt;color: #000000:#000000}
.news{FONT-SIZE:10pt; LINE-HEIGHT:14pt; color; color: #000000:#000000}
.contentbig{FONT-SIZE:11pt; LINE-HEIGHT:14pt;}
.info{  font-size: 9pt; line-height: 9pt;  color: #FFFFFF}
.footer{  font-size: 9pt; line-height: 12pt; font-weight: normal}
.search {  font-size: 10pt; line-height: 14pt; color: #ffffff; background-color: #75AEE3}
.whitehead {  font-size: 12pt; line-height: 15pt; color: #FFFFFF}
.whitecontent {  font-size: 10pt; line-height: 14pt; color: #ffffff}
.bgcolor {  background-color: #006797}
.leftline {  background-color: #FD7D04}
a:active {  color: #000000;; text-decoration: none}
a:visited {  color: #000000; font-weight: normal;; text-decoration: none}
a:link {  color: #000000; font-weight: normal; ; text-decoration: none}
a.homepage:link {  color: #000000; font-weight: normal;}
a.homepage:visited {  color: #000000; font-weight: normal;}
a.homepage:active {  color: #000000; font-weight: normal;}
a.homepage:hover {  color: #000000; font-weight: normal;}
</style>
</head>

<%
   sn=trim(Request.QueryString("sn"))
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select * FROM cwys_km where sn='"&sn&"'"
   obj.Source=sql
   obj.Open

%>

<body>


<form method="post" action="km_edit1.asp?sn=<%=sn%>" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="350" border="0">

<tr height="20" bgcolor="#E3E99E" align="left">
<td colspan=2 align=center> <font class=px14 color=blue>科目修改表单</font></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="left">
<td> <font class=px12>年份：<INPUT readonly value="<%=obj("nian")%>" type="text"  name=nian size=8></td>
<td> <font class=px12>部门：<INPUT readonly value="<%=obj("depar")%>" type="text"  name=depar></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=2> <font class=px12>一、二级科目代码：<INPUT value="<%=obj("kmcode")%>" type="text"  name=kmcode size=8></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=2> <font class=px12>一、二级科目说明：<INPUT value="<%=obj("kmshuom")%>" type="text"  name=kmshuom size=30></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=2> <font class=px12>三、四级科目代码：<INPUT value="<%=obj("fkmcode")%>" type="text"  name=fkmcode size=8></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=2> <font class=px12>三、四级科目说明：<INPUT value="<%=obj("fkmshuom")%>" type="text"  name=fkmshuom size=30></td>
</tr>
<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=2> <font class=px12>修改原因：<INPUT  type="text"  name=beizhu size=40></td>
</tr>

<tr height="20" bgcolor="#E3E9EE" align="center">
<td colspan=2> <input value="完成" type="submit" id="submit1" name="submit1"> </td>
</tr>
</table>
</form>
</body>
</html>
