<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage1.ClsUserManage1")
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
   nian=trim(Request.QueryString("nian"))
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select * FROM cwys_ed where sn='"&sn&"'"
   obj.Source=sql
   obj.Open
   

%>
   <%  
   Set obj1=server.CreateObject ("ADODB.Recordset")
   obj1.LockType=3
   obj1.CursorType=3
   set obj1.activeConnection=objConn
   sql1="select distinct * from cwys_km where  fkmcode='"&obj("fkmcode")&"' and nian='"&nian&"' "
   obj1.Source=sql1
   obj1.Open
   %>
   <% 
   if not obj1.eof then
   fkmshuom=obj1("kmshuom")+obj1("fkmshuom")
   else
   fkmshuom=obj("fkmcode")
   end if
   %>
<body>


<form method="post" action="ys_edit1.asp?sn=<%=sn%>" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="550" border="0">

<tr height="30" bgcolor="#E3E99E" align="left">
<td colspan=4 align=center> <font class=px14 color=blue>预算额度修改表单</font></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>年份：<%=obj("ys_year")%> <input type=hidden name=nian value=<%=obj("ys_year")%>></td>
<td width=25%> <font class=px12>部门：<%=obj("depar")%><input type=hidden name=depar value=<%=obj("depar")%>></td>
<td width=25%> <font class=px12>科目：<%=fkmshuom%><input type=hidden name=kemu value=<%=fkmshuom%>></td>
<td width=25%> <font class=px12>年度预算：<br><INPUT value="<%=obj("niandu")%>" type="text"  name=niandu size=12 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>一月：<INPUT value="<%=obj("jan")%>" type="text"  name=jan size=10  onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>二月：<INPUT value="<%=obj("feb")%>" type="text"  name=feb size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>三&nbsp;&nbsp月：<INPUT value="<%=obj("mar")%>" type="text"  name=mar size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>四&nbsp;&nbsp月：<INPUT value="<%=obj("apr")%>" type="text"  name=apr size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>五月：<INPUT value="<%=obj("may")%>" type="text"  name=may size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>六月：<INPUT value="<%=obj("jun")%>" type="text"  name=jun size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>七&nbsp;&nbsp月：<INPUT value="<%=obj("jul")%>" type="text"  name=jul size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>八&nbsp;&nbsp月：<INPUT value="<%=obj("aug")%>" type="text"  name=aug size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td > <font class=px12>九月：<INPUT value="<%=obj("sep")%>" type="text"  name=sep size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十月：<INPUT value="<%=obj("oct")%>" type="text"  name="oct" size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十一月：<INPUT value="<%=obj("nov")%>" type="text"  name=nov size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十二月：<INPUT value="<%=obj("dece")%>" type="text"  name=dece size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>
<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=4> <font class=px12>是否可超支：
<%
if obj("isover") then
isover="是"
else
isover="否"
end if
%>
    <select name="isover" style="HEIGHT: 22px; WIDTH: 57px"> 
    <option selected value="<%=obj("isover")%>"><%=isover%></option>
    <option value="True">是</option>
    <option value="False">否</option>
    </select> </td>
</tr>
<tr height="20" bgcolor="#E3E9EE" align="left">
<td colspan=4 align=left><font class=px12 color=red>备注：<INPUT  type="text"  name=beizhu size=60>(写入修改原因！)</font></td>
</tr>
<% 
obj.close
obj1.close
%>
<tr height="20" bgcolor="#E3E9EE" align="center">
<td colspan=4> <input value="完成" type="submit" id="submit1" name="submit1"> </td>
</tr>



<script language="JavaScript">
<!--
function display()
{
//document.writeln("kadsldsllkdadlklkasdlk")
//document.
form1.text1.value=Math.abs(form1.niandu.value-form1.jan.value-form1.feb.value-form1.mar.value-form1.apr.value-form1.may.value-form1.jun.value-form1.jul.value-form1.aug.value-form1.sep.value-form1.oct.value-form1.nov.value-form1.dece.value).toFixed(3)
//form1.text1.display=true
}
//-->
</script>



</table>
<br>
<br>
<p align="center">
<font color=blue>年度总额与各月累加和之差为：</font><INPUT  style="background-color: white; color: red;class=px14; border: 0 solid #9E9E9E"  readonly type="text" name=text1 id=text1 size=20>
</p>

</form>
</body>
</html>

