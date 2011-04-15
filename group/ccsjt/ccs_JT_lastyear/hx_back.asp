<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"ccs_gscn")
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
<title>深圳公司预算管理系统</title>
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
<script language="JavaScript">
<!--
function reload()
{
form1.submit();
window.opener.document.location.reload();
//window
window.close();
}
//-->
</script>
</head>

<%sub add_form%>
<%
   sn=trim(Request.QueryString("sn"))
%>

   <%  
  'Set obj1=server.CreateObject ("ADODB.Recordset")
   'obj1.LockType=3
   'obj1.CursorType=3
   'set obj1.activeConnection=objConn
   'sql1="select distinct * from cwys_km where  fkmcode='"&obj("fkmcode")&"' and nian='"&nian&"' "
   'obj1.Source=sql1
   'obj1.Open
   %>
   <% 
   'if not obj1.eof then
   'fkmshuom=obj1("fkmshuom")
   'else
   'fkmshuom=obj("fkmcode")
  ' end if
   %>
<body>


<form method="post" action="hx_back.asp?todo=01&sn=<%=sn%>" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="550" border="0">

<tr height="30"  align="left">
<td colspan=3 align=center> <font class=px14 color=blue>退回报销表单</font></td>
</tr>

<tr height="20" align="center">
<td colspan=4> <input style="color: red" value="确定退回" type="submit" name="action" onclick="javascript:reload()"> </td>
</tr>
</table>
</form>

<%end sub%>

<%sub save_data1%>
<%
sn=trim(Request.QueryString("sn"))
%>

<%
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("OledbStr") 
      
Set obj=server.CreateObject ("ADODB.Recordset")
obj.LockType=3
obj.CursorType=3
set obj.activeConnection=objConn
sql="select * FROM cwys_infoin where record_id='"&sn&"'"
obj.Source=sql
obj.Open
obj("ifhandin")="否"
                                                                                                                      
obj.Update                                                                          
obj.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>核销成功。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
</body>
</html>
<%end sub%>

<%'主过程                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                        
         add_form()   
              
       case "01"                                                                         
         save_data1()                                        
End Select                                                                       
%>