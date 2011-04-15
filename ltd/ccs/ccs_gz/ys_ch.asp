<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"ccs_cwytts")
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
window.close();
}
//-->
</script>
</head>

<%sub add_form%>
<%
    kemu=trim(Request.QueryString("kemu"))
    yue=trim(Request.QueryString("mnytime"))
    nian=trim(Request.QueryString("nian"))
    depar=trim(session("depar"))
    
    mnytime=cdate(nian+"-"+yue+"-1")
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
%>
   <%  
   Set obj1=server.CreateObject ("ADODB.Recordset")
   obj1.LockType=3
   obj1.CursorType=3
   set obj1.activeConnection=objConn
   sql1="select * from cwys_infoin where  mnydepm='"&depar&"'  and mnyyear='"&nian&"' and mnykmcode='"&kemu&"' and ((mnytime='"&mnytime&"' and  payway='预提') or (fromtime='"&mnytime&"' and payway='冲销')) order by payway desc,djdate desc "
   obj1.Source=sql1
   obj1.Open
      
   Set obj2=server.CreateObject ("ADODB.Recordset")
   obj2.LockType=3
   obj2.CursorType=3
   set obj2.activeConnection=objConn
   sql2="select  sum(price) as shengyu from cwys_infoin where  mnydepm='"&depar&"' and mnyyear='"&nian&"' and mnykmcode='"&kemu&"' and ((mnytime='"&mnytime&"' and  payway='预提') or (fromtime='"&mnytime&"' and payway='冲销'))  "
   obj2.Source=sql2
   obj2.Open
   shengyu=obj2("shengyu")
   obj2.Close
   %>
<form method="post" action="ys_ch.asp?todo=01&shengyu=<%=shengyu%>"  id="form1" name="form1">
   <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 0px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="700" border="1">      
     <tr><td colspan=7 align=center> <font class="px14" color="blue">预提、冲销记录</font></td></tr>
     <tr height="20" bgcolor="#E3E99E" align="left">
     <td align=center><font class="px12" color="black">部门</font></td>
     <td ><font class="px12" color="black">入帐月份</font></td>
     <td ><font class="px12" color="black">科目名称</font></td>
     <td ><font class="px12" color="black">金额</font></td>
     <td ><font class="px12" color="black">摘要</font></td>
     <td ><font class="px12" color="black">操作</font></td>
     <td ><font class="px12" color="black">操作日期</font></td>
     </tr>
   <%
   obj1.MoveFirst 
   while not obj1.EOF
   %> 
     <tr height="20" bgcolor="#E3E9EE" align="left">
     <td ><font class="px12" color="black"><%=obj1("mnydepm")%></font></td>
     <td ><font class="px12" color="black"><%=year(obj1("mnytime"))%>年<%=month(obj1("mnytime"))%>月</font></td>
     <td ><font class="px12" color="black"><%=obj1("mnykm")%></font></td>
     <td ><font class="px12" color="black"><%=obj1("price")%></font></td>
     <td ><font class="px12" color="black"><%=obj1("mnynote")%></font></td>
     <td ><font class="px12" color="black"><%=obj1("payway")%></font></td>
     <td ><font class="px12" color="black"><%=obj1("djdate")%></font></td>
     </tr> 

   <%
   mnykm=obj1("mnykm")
   obj1.movenext
   wend
   obj1.Close
   %>	
   <tr height="20"><td colspan=7 align=center> <font class="px12" color="red">剩余金额：<%=shengyu%>元</font></td></tr>
      <tr><td colspan=7 align=center> <font class="px12" color="red">冲销金额：<INPUT type="text"  name=jine size=8>元(输入正数即可)&nbsp;&nbsp;&nbsp;&nbsp;
      充入区间：
          
    <select name="nian" style="HEIGHT: 22px; WIDTH: 50px"> 
    <option selected><%=year(date)%></option>
    <option><%=year(date)-1%></option>
    <option><%=(year(date)+1)%></option>
    </select>年

    <select name="yue" style="HEIGHT: 22px; WIDTH: 50px"> 
    <option selected><%=month(date)%></option>
        <%       
          yue=1
          while yue<=12 
        %>
        <option><%=yue%></option>
        <%
        yue=yue+1
        wend%></select>月</td></tr>
<tr height="20">
 <td colspan=7 align="center"> 
    <font class="px12" color="red">摘要：<INPUT  type="text"  name="note" size=80></font>
    <input value="确定" type="submit" name="action" ></font>
 </td>
</tr>
  </table>
  <input type="hidden" name=kemu value=<%=kemu%>>
  <input type="hidden" name=mnytime value=<%=mnytime%>>
  <input type="hidden" name=depar value=<%=depar%>>
  <input type="hidden" name=mnykm value=<%=mnykm%>>
</form>  
  
<body>



<%end sub%>

<%sub save_data1%>

<%

      kemu=trim(Request.form("kemu"))
      depar=trim(Request.form("depar"))
      mnytime=trim(Request.form("mnytime"))
      mnykm=trim(Request.form("mnykm"))
      
shengyu=trim(Request.QueryString("shengyu"))
jine=Request.form("jine")
note=trim(Request.form("note"))
nian=trim(Request.form("nian"))
yue=trim(Request.form("yue"))
fromtime=cdate(nian+"-"+yue+"-1")
%>


<%if jine="" then%>
<table width=700>
<tr align="center" width=700>
<td align="center" width=700>
	<font class=px14 color=blue>输入金额为空，请重新输入。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:history.go(-1)">
</td>
</tr>
</table>
<%elseif note="" then%>
<table width=700>
<tr align="center" width=700>
<td align="center" width=700>
	<font class=px14 color=blue>摘要为空，请输入摘要内容。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:history.go(-1)">
</td>
</tr>
</table>
<%elseif clng(shengyu)-clng(jine)<0 then%>
<table width=700>
<tr align="center" width=700>
<td align="center" width=700>
	<font class=px14 color=blue>所输的金额大于剩余额，请重新输入。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
</table>
<%else%>


<%
   Set objConn1 = Server.CreateObject("ADODB.Connection")
   objConn1.Open Application("OledbStr")
   Set obj1=server.CreateObject ("ADODB.Recordset")
   obj1.LockType=3
   obj1.CursorType=3
   set obj1.activeConnection=objConn1
   sql="select * from cwys_infoin "
   obj1.Source=sql
   obj1.Open
%> 

<%
objConn1.BeginTrans                                                                          
obj1.AddNew
obj1("passname")="财务"
obj1("djdate")=date
obj1("djname")=trim(session("emid"))
obj1("hxdate")=date
obj1("hxname")=trim(session("emid"))
obj1("mnydepm")=depar
obj1("mnykmcode")=kemu
obj1("mnykm")=mnykm
obj1("mnynote")=note
obj1("price")=0-jine
obj1("payway")="冲销"
obj1("ifhx")="是"
obj1("ifhandin")="是"
obj1("mnytime")=fromtime
obj1("fromtime")=mnytime
obj1("mnyyear")=nian
obj1("cz")="录入"
                             
obj1.Update                                                                          
objConn1.CommitTrans 
obj1.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>修改成功。</font></p>
	<input type="button" name="button" value="确定" onclick="javascript:reload()">
</td>
</tr>
</table>
</body>
</html>
<%end if%>
<%end sub%>

<%'主过程                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                        
         add_form()   
              
       case "01"                                                                         
         save_data1()                                        
End Select                                                                       
%>