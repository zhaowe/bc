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
<body>
<% Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn%>
<%q=Request.QueryString("q")

km1=Request.QueryString ("km1")
session("km1")=km1
km2=Request.QueryString ("km2")
session("km2")=km2
km=Request.QueryString ("km")
session("km")=km
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")   
  
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
'Response.Write(sqlb)
rs_r3.Open sqlr,conn_1,3,3,1 
  
  
  year1=Request.QueryString("year1")
  year2=Request.QueryString("year2")
  month1=Request.QueryString("month1")
  month2=Request.QueryString("month2")
objrst.Source ="delete from cwys_infoin where record_id='"& q &"'"
objrst.Open 



%>

<font class="px12"><%=q%>号记录已经被删除</font>
<%Response.Redirect ("ccs_bm_ser.asp?cz=删除")
%>

</BODY>
</HTML>