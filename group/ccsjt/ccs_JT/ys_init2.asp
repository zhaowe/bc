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

window.opener.history.go(-1);
//window

}
//-->
</script>
</head>
<%
   nian=trim(Request.QueryString("nian"))
   depar=trim(Request.QueryString("depar"))
   kemu=trim(Request.QueryString("kemu"))
   edu=trim(Request.QueryString("edu"))
%>
<%sub add_form()%>
<body>

<form method="post" action="ys_init2.asp?todo=01" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="550" border="0">

<tr height="30" bgcolor="#E3E99E" align="left">
<td colspan=4 align=center> <font class=px14 color=blue>预算额度分配表单</font></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>年份：<%=nian%> <input type=hidden name=nian value=<%=nian%>></td>
<td width=25%> <font class=px12>部门：<%=depar%><input type=hidden name=depar value=<%=depar%>></td>
<td width=25%> <font class=px12>科目：<%=kemu%><input type=hidden name=kemu value=<%=kemu%>></td>
<td width=25%> <font class=px12>年度预算：<br><INPUT readonly value="<%=edu%>" type="text"  name=niandu size=16 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>一月：<INPUT type="text"  name=jan size=10  onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>二月：<INPUT type="text"  name=feb size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>三&nbsp;&nbsp月：<INPUT type="text"  name=mar size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>四&nbsp;&nbsp月：<INPUT type="text"  name=apr size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td width=25%> <font class=px12>五月：<INPUT  type="text"  name=may size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>六月：<INPUT  type="text"  name=jun size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>七&nbsp;&nbsp月：<INPUT  type="text"  name=jul size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td width=25%> <font class=px12>八&nbsp;&nbsp月：<INPUT  type="text"  name=aug size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td > <font class=px12>九月：<INPUT  type="text"  name=sep size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十月：<INPUT  type="text"  name="oct" size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十一月：<INPUT type="text"  name=nov size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
<td > <font class=px12>十二月：<INPUT type="text"  name=dece size=10 onkeyup="javascript:display()" onchange="javascript:display()"></td>
</tr>

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
<%end sub%>
<%sub save_data1%>
<body onload="javascript:reload()">
<%
depar=trim(Request.form("depar"))
nian=trim(Request.form("nian"))
kemu=trim(Request.form("kemu"))
niandu=trim(Request.form("niandu"))
jan=trim(Request.form("jan"))
if jan="" then 
jan=0
end if
feb=trim(Request.form("feb"))
if feb="" then
	feb=0
end if
mar=trim(Request.form("mar"))
if mar="" then
	mar=0
end if
apr=trim(Request.form("apr"))
if apr="" then
	apr=0
end if
may=trim(Request.form("may"))
if may="" then
	may=0
end if
jun=trim(Request.form("jun"))
if jun="" then
	jun=0
end if
jul=trim(Request.form("jul"))
if jul="" then
	jul=0
end if
aug=trim(Request.form("aug"))
if aug="" then
	aug=0
end if
sep=trim(Request.form("sep"))
if sep="" then
	sep=0
end if
shiyiy=trim(Request.form("oct"))
if shiyiy="" then
	shiyiy=0
end if
nov=trim(Request.form("nov"))
if nov="" then
	nov=0
end if
dece=trim(Request.form("dece"))
if dece="" then
	dece=0
end if
%>
<%
dd=abs(cdbl(niandu)-cdbl(jan)-cdbl(feb)-cdbl(mar)-cdbl(apr)-cdbl(may)-cdbl(jun)-cdbl(jul)-cdbl(aug)-cdbl(sep)-cdbl(shiyiy)-cdbl(nov)-cdbl(dece))
''if clng(niandu)<>clng(jan)+clng(feb)+clng(mar)+clng(apr)+clng(may)+clng(jun)+clng(jul)+clng(aug)+clng(sep)+clng(shiyiy)+clng(nov)+clng(dece) then
if dd>0.001  then
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>各月预算额度的累加值与年度总额不一致，请返回检查。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
</table>
<%else%>
<%
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("OledbStr") 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("OledbStr") 
   
Set obj1=server.CreateObject ("ADODB.Recordset")
obj1.LockType=3
obj1.CursorType=3
set obj1.activeConnection=objConn1
sql="select * from cwys_ed "
obj1.Source=sql
obj1.Open
%>
<%'将操作写入日志
descr=descr+"初始化"+depar+nian+"年"+kemu+"年度总额为:"+cstr(niandu)+"一月份额度为："+cstr(jan)+"二月份额度为："+cstr(feb)+"三月份额度为："+cstr(mar)
descr=descr+"四月份额度为："+cstr(apr)+"五月份额度为："+cstr(may)+"六月份额度为："+cstr(jun)+"七月份额度为："+cstr(jul)+"八月份额度为："+cstr(aug)+"九月份额度为："+cstr(sep)+"十月份额度为："+cstr(shiyiy)+"十一月份额度为："+cstr(nov)+"十二月份额度为："+cstr(dece)
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2  
sqlb="insert into cwys_log (operation,descr,type,operatetime,operator) values ('初始化','"&descr&"','预算','"&now&"','"&session("emid")&"')" 
rs_b.Open sqlb,objConn,3,3,1 
'rs_b.Close

%>

<%
objConn1.BeginTrans                                                                          
obj1.AddNew
obj1("ys_year")=nian
obj1("depar")=depar
obj1("fkmcode")=kemu
obj1("niandu")=niandu
obj1("jan")=jan
obj1("feb")=feb
obj1("mar")=mar
obj1("apr")=apr
obj1("may")=may
obj1("jun")=jun
obj1("jul")=jul
obj1("aug")=aug
obj1("sep")=sep
obj1("oct")=shiyiy
obj1("nov")=nov
obj1("dece")=dece
obj1("isover")=0


obj1("lururen")=trim(session("emid"))
obj1("lurudate")=date                                                                   
                                                     
obj1.Update                                                                          
objConn1.CommitTrans 
obj1.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>录入成功。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
<%end if%>
</body>

<%end sub%>
<%'主过程                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                        
         add_form()   
              
       case "01"                                                                         
         save_data1()
         'Response.redirect("TMS_plan.asp")                                        
End Select                                                                       
%>

