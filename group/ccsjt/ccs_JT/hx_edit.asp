<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage1.ClsUserManage1")
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
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select * FROM cwys_infoin where record_id='"&sn&"'"
   obj.Source=sql
   obj.Open
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


<form method="post" action="hx_edit.asp?todo=01&sn=<%=sn%>" id="form1" name="form1">
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="550" border="0">

<tr height="30" bgcolor="#E3E99E" align="left">
<td colspan=3 align=center> <font class=px14 color=blue>报销帐目核销表单</font></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td> <font class=px12>经办人员工号：<%=obj("passcode")%></td>
<td> <font class=px12>经办人姓名：<%=obj("passname")%></td>
<td> <font class=px12>费用控制部门：<%=obj("mnydepm")%></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td > <font class=px12>报销人员工号：<%=obj("bxcode")%></td>
<td > <font class=px12>报销人姓名：<%=obj("bxname")%></td>
<td > <font class=px12>费用科目：

    <%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open Application("OledbStr") 
    Set objRst=server.CreateObject ("ADODB.Recordset")
    objRst.LockType=3
    objRst.CursorType=3
    set objRst.activeConnection=Conn
    
    objrst.Source ="select * from cwys_km where depar = '"&obj("mnydepm")&"' and nian='"&obj("mnyyear")&"' order by sn" 
    objrst.Open 
    %>
    
<select name="fkm">
         <option selected value="<%=obj("mnykmcode")%>@<%=obj("mnykm")%>"><%=obj("mnykm")%></option>
      <%while not objrst.EOF %>
    <option value="<%=objrst("fkmcode")%>@<%=objrst("fkmshuom")%>"><%=objrst("fkmshuom")%></option>
    <%objrst.MoveNext %> 
    <% wend %>
    <%objrst.Close%>
</select>   
</td>     
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">

<td > <font class=px12>费用期间：<INPUT readonly value="<%=obj("mnyyear")%>" type="text"  name=nian size=4>年<INPUT value="<%=month(obj("mnytime"))%>" type="text"  name=yue size=4>月</td>
<td > <font class=px12>付款方式：<select name="payway">
         <OPTION selected><%=obj("payway")%></OPTION>
         <OPTION  value="现金"><font color=black class="px14">现金</OPTION>
         <OPTION  value="银行"><font color=black class="px14">银行</OPTION>
         <OPTION  value="内部往来"><font color=black class="px14">内部往来</OPTION>
         </select></td>
<td > <font class=px12>金额：<INPUT value="<%=obj("price")%>" type="text"  name=jine size=8></td>
</tr>

<tr height="25" bgcolor="#E3E9EE" align="left">
<td colspan=3> <font class=px12>费用说明：<INPUT value="<%=obj("mnynote")%>" type="text"  name=fynote size=60>
</td>

</tr>

<% obj.close%>
<tr height="20" bgcolor="#E3E9EE" align="center">
<td colspan=4> <input value="确定" type="submit" name="action" onclick="javascript:reload()"> </td>
</tr>
</table>
</form>

<%end sub%>

<%sub save_data1%>
<%
sn=trim(Request.QueryString("sn"))
nian=Request.form("nian")
yue=Request.form("yue")
mnytime=cdate(nian+"-"+yue+"-1")
payway=trim(Request.form("payway"))
price=Request.form("jine")
mnynote=trim(Request.form("fynote"))
fkm=trim(Request.form("fkm"))
Dim ar1
ar1=Split(fkm,"@")
mnykm=ar1(1)
mnykmcode=ar1(0)
fkmshuom=mnykm


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
%>
<!--
<%     
Set obj0=server.CreateObject ("ADODB.Recordset")
obj0.LockType=3
obj0.CursorType=3
set obj0.activeConnection=objConn
sql0="select * FROM cwys_km where fkmshuom='"&mnykm&"' and depar='"&obj("mnydepm")&"' and nian='"&obj("mnyyear")&"'"
obj0.Source=sql0
obj0.Open
   if not obj0.eof then
   fkmshuom=trim(obj0("fkmshuom"))
   else
   fkmshuom=""
   end if
   obj0.Close
%>
-->

<%    
kytime=cdate(nian+"-"+yue+"-1")
  
Set obj1=server.CreateObject ("ADODB.Recordset")
obj1.LockType=3
obj1.CursorType=3
set obj1.activeConnection=objConn
sql1="select ISNULL(SUM(price), 0) FROM cwys_infoin where mnydepm='"&obj("mnydepm")&"' and mnytime<='"&kytime&"' and mnyyear='"&obj("mnyyear")&"' and mnykmcode='"&mnykmcode&"' and  cz<>'删除' and ifhandin='是' and record_id<>'"&sn&"' "
obj1.Source=sql1
obj1.Open
yiyong=obj1(0)
obj1.Close

Set obj2=server.CreateObject ("ADODB.Recordset")
obj2.LockType=3
obj2.CursorType=3
set obj2.activeConnection=objConn
sql2="select * FROM cwys_ed where depar='"&obj("mnydepm")&"' and ys_year='"&obj("mnyyear")&"'   and fkmcode='"&mnykmcode&"' "
obj2.Source=sql2
obj2.Open
if obj2.EOF then
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=red>还没有给<%=mnykm%>科目分配额度。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
<%else%>
<%

if yue=1 then
keyong=obj2("jan")
elseif yue=2 then
keyong=obj2("jan")+obj2("feb")
elseif yue=3 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")
elseif yue=4 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")
elseif yue=5 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")
elseif yue=6 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")
elseif yue=7 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")+obj2("jul")
elseif yue=8 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")+obj2("jul")+obj2("aug")
elseif yue=9 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")+obj2("jul")+obj2("aug")+obj2("sep")
elseif yue=10 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")+obj2("jul")+obj2("aug")+obj2("sep")+obj2("oct")
elseif yue=11 then
keyong=obj2("jan")+obj2("feb")+obj2("mar")+obj2("apr")+obj2("may")+obj2("jun")+obj2("jul")+obj2("aug")+obj2("sep")+obj2("oct")+obj2("nov")
else
keyong=obj2("niandu")
end if





%>

<%if keyong-yiyong-price<0 and trim(obj("isover"))="否" then%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=red>已超支，修改不成功。截止本月剩余<%=(keyong-yiyong)%>元,欲报销<%=price%>元。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
<%else%>
<%
obj("mnykm")=trim(mnykm)
obj("mnykmcode")=trim(mnykmcode)
obj("payway")=payway
obj("price")=price
obj("mnynote")=mnynote
obj("mnytime")=mnytime                                                                                                                     
obj.Update                                                                          
obj.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>修改成功。截止本季度剩余<%=(keyong-yiyong)%>元,报销<%=price%>元。现剩余<%=(keyong-yiyong-price)%>元。</font></p>
	<input type="button" name="button" value="确定" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
<%end if%>
<%end if%>
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