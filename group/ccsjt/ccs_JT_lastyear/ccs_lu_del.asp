<%@ Language=VBScript %>
<HTML>
<HEAD>
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
function check()
{
if (confirm("你确定要删除吗？")==false)
  return false
}

</script>

<body>
<% Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
    Set objRst1=server.CreateObject ("ADODB.Recordset")
  objRst1.LockType=3
  objRst1.CursorType=3
  set objRst1.activeConnection=objConn%>
  <% 
 

e=trim(session("emid"))
t=trim(session("loginid"))
'Response.Write(e)
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")                                                                           

set rs_1=server.CreateObject("adodb.recordset")                                                                            
rs_1.CursorLocation=2                                                                            
sql1="SELECT * FROM logininfo WHERE loginid='"& t &"'"                                                                            
rs_1.Open sql1,conn_1,3,3,1
if not rs_1.EOF then
cid=rs_1("companyid")
name=rs_1("name")
else
Response.Write("请重新登陆")
end if

set rs_3=server.CreateObject("adodb.recordset")
rs_3.CursorLocation=2
sql3="select * from companyinfo where companyid='"&cid&"'"
rs_3.Open sql3,conn_1,3,3,1

f=trim(rs_3("description"))
'Response.Write(f)%>


<%q=Request.QueryString("q")
  km=Request.QueryString("km")

km1=Request.QueryString ("km1")
  
 set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2  
sql="select * from cwys_infoin where record_id='"&q&"'"
rs_2.open sql,conn_1,3,3,1
passcode=trim(rs_2("passcode"))
passname=trim(rs_2("passname"))
bxcode=trim(rs_2("bxcode"))
bxname=trim(rs_2("bxname"))
djdate=trim(rs_2("djdate"))
djname=trim(rs_2("djname"))
mnydepm=trim(rs_2("mnydepm"))
mnykm=trim(rs_2("mnykm"))
mnynote=trim(rs_2("mnynote"))
price=trim(rs_2("price"))
date1=trim(rs_2("mnytime"))
payway=trim(rs_2("payway"))
ifhandin=trim(rs_2("ifhandin"))
mnyyear=trim(rs_2("mnyyear"))
 
 
 set rs_r3=server.CreateObject("adodb.recordset")                                                                            
rs_r3.CursorLocation=2                                                                            
sqlr="insert into cwys_bmglrz (passcode,passname,bxcode,bxname,djdate,djname,mnydepm,mnykm,mnynote,price,mnytime,payway,ifhandin,ifhx,mnyyear,changeid,cz,lururen,lurutime,mnykmcode) values ('"&passcode&"','"&passname&"','"&bxcode&"','"&bxname&"','"&djdate&"','"&djname&"','"&mnydepm&"','"&mnykm&"','"&mnynote&"','"&price&"','"&date1&"','"&payway&"','"&ifhandin&"','否','"&mnyyear&"','"&q&"','删除','"&session("emid")&"','"&date()&"','"&km1&"')"                                                                            
'Response.Write(sqlr)
rs_r3.Open sqlr,conn_1,3,3,1 
  

'objrst1.Source ="update cwys_infoin set cz='删除' where record_id='"& q &"'"
'objrst1.Open 

objrst1.Source ="delete from cwys_infoin where record_id='"& q &"'"
objrst1.Open
%>






<font class="px12" color="red"><%=q%>号记录已经被删除</font>
<% objrst.Source = "select * from cwys_infoin  where  mnykmcode='" & km1 &"' and mnydepm='"&f&"' and ifhx='否' and cz<>'删除' order by record_id desc"
   objRst.Open
   'Response.Write(objrst.source) 
   if objrst.EOF and objrst.BOF   then %>    
    <font color=black class=px12><STRONG>没有你要的<%=km%>的信息
   <%else%>
   
<table align="left"  cellSpacing="0" cellPadding="0" width="400" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="73%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="561" border="0">
        <tbody>
   
        <tr>
          <td vAlign="top" width="564">
            <table cellSpacing="1" cellPadding="0" width="700">
              <tbody>
               <tr bgColor="#9CD7F5" height="20" width="400">
               <td align="middle" width="77"  ><font color=black class=px12>操作</td>
                <td align="middle" width="84"  ><font color=black class=px12>帐目时间</td>
              
                <td align="middle" width="200"  ><font color=black class=px12>子科目</td>
                <td align="middle" width="42"  ><font color=black class=px12>金额</td>
                <td align="middle" width="84"  ><font color=black class=px12>是否提交</td>
                <td align="middle" width="84"  ><font color=black class=px12>简要说明</td>
                <td align="middle" width="64"  ><font color=black class=px12>报销人</td>
                <td align="middle" width="64"  ><font color=black class=px12>录入人</td>
                <td align="middle" width="64"  ><font color=black class=px12>记录号</td>
                <td align="middle" width="64"  ><font color=black class=px12>帐单号</td>
                
                <td align="middle" width="77"  ><font color=black class=px12>操作</td>
                </tr>
                <%
               
                do while not objrst.EOF%>
                <tr bgColor="#ecf7fd" height="20">
                <%if trim(objrst("ifhandin"))="否" then%>
                <td align="middle"  ><font color=blue class=px12 onClick="JavaScript:window.open('ccs_lu_xg2.asp?q=<%=objrst("record_id")%>','hh','width=550,left=200,top=10,height=242');">提交</font></a></td>
                <%else%>
                <td align="middle"  ><font color="#cccccc" class=px12>提交</font></a></td>
                 <%end if%>                <td align="middle" width="84" ><font color=black class=px12><%=year(objrst("mnytime"))%>年<%=month(objrst("mnytime"))%>月</td>
                <td align="middle" width="200" ><font color=black class=px12><%=objrst("mnykm")%></td>
                <td align="middle" width="42"  ><font color=black class=px12><%=objrst("price")%></td>

    <%if trim(objrst("ifhandin"))="否" then%>
                <td align="middle"  ><font color="#ff9900" class=px12><%=objrst("ifhandin")%></td>
                <%else%>
                <td align="middle"  ><font color=black class=px12><%=objrst("ifhandin")%></td>
                <%end if%>

                <td align="middle" width="84"  ><font color=black class=px12><%=objrst("mnynote")%></td>
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("bxname")%></td>
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("djname")%></td>
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("record_id")%></td>
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("tabid")%></td>
                <td align="middle" width="77"  ><a href="ccs_lu_del.asp?q=<%=objrst("record_id")%>&km=<%=objrst("mnykm")%>&km1=<%=km1%>" onclick="return check()"><font color=red class=px12>删除</font></a></td>
                </tr>
                <%
                objrst.MoveNext 
                loop%> 
          
             </tbody></table></td></tr></tbody></table>
<%end if%>

</table>
</strong></font>

</BODY>
</HTML>