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
form1.submit();
window.opener.document.location.reload();
//window
//window.close();
}
//-->
</script>
</head>

<%sub add_form%>
<%
   sn=trim(Request.QueryString("sn"))
%>

<body>


<form method="post" action="hx_del.asp?todo=01&sn=<%=sn%>" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="350" border="0">

<tr height="20"  >
<td align=center> <input value="确定删除" style="color: blue;" type="submit" name="action" onclick="javascript:reload()"> </td>
</tr>


</table>
</form>

<%end sub%>

<%sub save_data1%>
<%
   sn=trim(Request.QueryString("sn"))
   sn=cint(sn)
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   Set obj1=server.CreateObject ("ADODB.Recordset")
   obj1.LockType=3
   obj1.CursorType=3
   set obj1.activeConnection=objConn
   sql1="select * from cwys_infoin  where record_id='"&sn&"'"
   obj1.Source=sql1
   obj1.Open 
   hxdate=trim(obj1("hxdate"))
   hxname=trim(obj1("hxname"))
   bxname=trim(obj1("bxname"))
   mnydepm=trim(obj1("mnydepm"))
   mnykm=trim(obj1("mnykm"))
   mnynote=trim(obj1("mnynote"))
   price=trim(obj1("price"))
   obj1.Close
%>
<%'将操作写入日志


descr="删除"+hxname+"与"+hxdate+"核销的"+mnydepm+bxname+"已报销的"+mnykm+price+"("+mnynote+")"
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2  
sqlb="insert into cwys_log (operation,descr,type,operatetime,depar,operator) values ('删除','"&descr&"','核销','"&now&"','"&mnydepm&"','"&session("emid")&"')" 
rs_b.Open sqlb,objConn,3,3,1 
'rs_b.Close

%>
<%
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="DELETE FROM cwys_infoin  where record_id='"&sn&"'"
   obj.Source=sql
   obj.Open
%>

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