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
<script language="JavaScript">
<!--
function reload()
{

window.opener.document.location.reload();
//window

}
//-->
</script>
</head>
<html>
<body onload="javascript:reload()">
<%
sn=trim(Request.QueryString("sn"))
nian=trim(Request.form("nian"))
depar=trim(Request.form("depar"))
kmcode=trim(Request.form("kmcode"))
kmshuom=trim(Request.form("kmshuom"))
fkmcode=trim(Request.form("fkmcode"))
fkmshuom=trim(Request.form("fkmshuom"))
beizhu=trim(Request.form("beizhu"))

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("OledbStr") 



      
Set obj=server.CreateObject ("ADODB.Recordset")
obj.LockType=3
obj.CursorType=3
set obj.activeConnection=objConn
sql="select * FROM cwys_km where sn='"&sn&"'"
obj.Source=sql
obj.Open


descr=descr+"修改"+depar+nian+"年"+fkmshuom+"科目，"
if trim(obj("kmcode"))<>kmcode then 
descr="一级科目代码由："+trim(obj("kmcode"))+"修改为："+kmcode+"；"
end if

if trim(obj("kmshuom"))<>kmshuom then
descr=descr+"一级科目说明由："+trim(obj("kmshuom"))+"修改为："+kmshuom+"；"
end if

if trim(obj("fkmcode"))<>fkmcode then
descr=descr+"二级科目代码由："+trim(obj("fkmcode"))+"修改为："+fkmcode+"；"
end if

if trim(obj("fkmshuom"))<>fkmshuom then
descr=descr+"二级科目说明由："+trim(obj("fkmshuom"))+"修改为："+fkmshuom+"；"
end if



if beizhu<>"" then
descr=descr+"修改原因为："+beizhu+"。"
else
descr=descr+"未填写修改原因。"
end if

set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2  
sqlb="insert into cwys_log (operation,descr,type,operatetime,operator) values ('修改','"&descr&"','科目','"&now&"','"&session("emid")&"')" 
rs_b.Open sqlb,objConn,3,3,1 

obj("kmcode")=kmcode
obj("kmshuom")=kmshuom  
obj("fkmcode")=fkmcode
obj("fkmshuom")=fkmshuom
'obj("isover")=isover
                                                                                                                      
obj.Update                                                                          
obj.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>修改成功。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
</body>
</html>
