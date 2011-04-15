
<%@ Language=VBScript %>


  
<html>
<head>
<title>公司预算管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<%q=Request.QueryString("q")
mnykm=Request.QueryString("mnykm")
date1=Request.QueryString("date1")
date2=Request.QueryString("date2")
'Response.Write(date1)
'Response.Write(date2)
'Response.Write(mnykm)
%>
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
'Response.Write("请重新登陆")
end if

set rs_3=server.CreateObject("adodb.recordset")
rs_3.CursorLocation=2
sql3="select * from companylocale where companyid='"&cid&"'"
rs_3.Open sql3,conn_1,3,3,1

f=trim(rs_3("companyname"))


'Response.Write(f)
'Response.Write(t)
km=Request.QueryString ("km")
' dep="货运部"
' session("dep")=dep
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
    Set objRst1=server.CreateObject ("ADODB.Recordset")
  objRst1.LockType=3
  objRst1.CursorType=3
  set objRst1.activeConnection=objConn
  
    Set objRst2=server.CreateObject ("ADODB.Recordset")
  objRst2.LockType=3
  objRst2.CursorType=3
  set objRst2.activeConnection=objConn
  
set rs_9=server.CreateObject("adodb.recordset")                                                                            
rs_9.CursorLocation=2                                                                            
'sql1="SELECT * FROM personinfo WHERE pid='"& e &"'"
sql1="SELECT distinct pid,pname FROM personinfo "                                                                                                                                                        
rs_9.Open sql1,conn_1,3,3,1  
  
set rs_8=server.CreateObject("adodb.recordset")                                                                            
rs_8.CursorLocation=2                                                                            
'sql1="SELECT * FROM personinfo WHERE pid='"& e &"'"
sql1="SELECT distinct pid,pname FROM personinfo "                                                                                                                                                        
rs_8.Open sql1,conn_1,3,3,1    
  %>
  


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
//form1.submit();
window.opener.document.location.reload();

window.close();
}

function MM_preloadimagess() { //v3.0
  var d=document; if(d.imagess){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadimagess.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new images; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapimages() { //v3.0
  var i,j=0,x,a=MM_swapimages.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}



function disname()
{
var ahbh=new Array;   
var ahd=new Array;      
var hds=new String;   
var hbhs;



<%
k=0
while not rs_9.EOF 
%>

ahbh[<% =k %>]='<%=trim(rs_9("pid"))%>'  ;        
ahd[<% =k %>]='<%=trim(rs_9("pname"))%>'  ; 
//window.alert(ahbh[0]);

<%          
k=k+1          
rs_9.MoveNext      
wend     
rs_9.close     
%> 


hbhs=document.form1.hbh.value;     
//   document.hdtj.hbh.value       
//hbh=hbh.toString();      
hds="";   
for(i=0;i<<%=k-1%>;i++)   
{        
if(hbhs==ahbh[i])            
hds=ahd[i];   
}  

document.form1.hd.value =hds;

}


function disname1()
{
var ahbh1=new Array;   
var ahd1=new Array;      
var hds1=new String;   
var hbhs1;

<%
k=0
while not rs_8.EOF 
%>

ahbh1[<% =k %>]='<%=trim(rs_8("pid"))%>'  ;        
ahd1[<% =k %>]='<%=trim(rs_8("pname"))%>'  ; 
//window.alert(ahbh[0]);

<%          
k=k+1          
rs_8.MoveNext      
wend     
rs_8.close     
%> 

hbhs1=document.form1.hbh1.value;

hds1="";   
for(i=0;i<<%=k-1%>;i++)   
{        
if(hbhs1==ahbh1[i])            
hds1=ahd1[i]; 

}
document.form1.hd1.value =hds1;  


}
//-->
</script>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333FF" onLoad="MM_preloadimagess('images/bt_01_off.gif','images/bt_02_on.gif','images/bt_03_on.gif','images/bt_04_on.gif','images/bt_05_on.gif')">
<%objrst1.Source ="select * from cwys_infoin where record_id = '"&q&"'" 
  objrst1.Open
   'Response.Write(objrst.Source)
  objrst2.Source="SELECT fkmcode FROM cwys_km WHERE fkmshuom='"& objrst1("mnykm") &"' and depar='"&f&"' and nian='"&year(date())&"'"
  objrst2.open 
  if not objrst2.EOF then
fkmcode2=objrst2("fkmcode")

else
fkmcode2="无"
end if                   
                      %>
                      



<%sub list()
  %>

<table width="740" border="0" cellspacing="0" cellpadding="0" height="90%">
  <tr>
    <td width="595" valign="top"> 
    <form method="post" action="ccs_lu_xg2.asp?todo=02&q=<%=q%>" id=form1 name=form1>   

     <table style="BORDER-RIGHT: #4e4c71 1px solid; BORDER-TOP: #4e4c71 1px solid; BORDER-LEFT: #4e4c71 1px solid; BORDER-BOTTOM: #4e4c71 1px solid" height="92%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#cccccc" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="black" class="px12">&nbsp;超标账单提交</font> </td></tr>
        <tr>
          <td vAlign="top" width="100%">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
              
              <tr>
               <%objrst.Source ="select * from cwys_infoin where record_id = '"&q&"'" 
                       objrst.Open
                       'Response.Write(objrst.Source)
                      %>
    <td align="right" >　</td></tr>
   <tr>
    <td align="center" ><font class="px12" color="black">您目前进入的是<font color=red><%=q%></font>号纪录<font color=red><%=objrst("mnykm")%></font>提交界面</td></tr>
    <tr>
    <td align="right" >　</td></tr>

                       
<tr width=100% bgcolor="white">
     <td  align="left" bgcolor="eeeeee">
     <font class="px12" color="black">&nbsp;&nbsp;经办人员工号：</font>
     <input id="hbh" type="text" name="hbh" size="5" onchange="javascript:disname()" value=<%=objrst("passcode")%> readonly>    
    
     <font class="px12" color="black">&nbsp;&nbsp;经办人姓名：</font>
     <input id="hd" type="text" name="hd" size="10" value=<%=objrst("passname")%> readonly>
     
     <font class="px12" color="black">&nbsp;&nbsp;费用控制部门：</font>
     <input type="text" name="mnydepm" size="10" readonly value=<%=f%>></td>
</tr>
 <tr bgcolor="white"><td bgcolor="eeeeee">
     <font class="px12" color="black">&nbsp;&nbsp;报销人员工号：</font>
     <input type="text" name="hbh1" size="5" onchange="javascript:disname1()" value=<%=objrst("bxcode")%> readonly> 
     
     <font class="px12" color="black">&nbsp;&nbsp;报销人姓名：</font>
     <input type="text" name="hd1" size="10" value=<%=objrst("bxname")%> readonly>
     
     <font class="px12" color="black">&nbsp;&nbsp;费用科目：</font>
     <input type="text" name="mnykm" size="20" readonly value=<%=objrst("mnykm")%> >

<font class="px12"><%=objrst("mnykmcode")%></font>  
 <input type="hidden" name="km1" size="16" readonly value=<%=objrst("mnykmcode")%>>    
<br>
   
      </td>
    </tr>


    <tr bgcolor="white">
      <td align="left" bgcolor="eeeeee">
    <font class="px12" color="black">&nbsp;&nbsp;费用期间：</font>
        <select size="1" name="year1"><option selected><%=objrst("mnyyear")%>
       
            </select>
            <font color=black class="px14">年</font>
        <select name="month1" style="HEIGHT: 22px; WIDTH: 43px" ><option selected> <%=month(objrst("mnytime"))%>
    
        </select>
      <font class="px12" color="black">月</font>
     
     <font class="px12" color="black">&nbsp;&nbsp;付款方式：</font>
        <select size="1" name="payway"><option selected><%=objrst("payway")%>
    

            </select>
                
     <font class="px12" color="black"><br>&nbsp;&nbsp;金额：</font>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="price" size="10" value=<%=objrst("price")%> readonly><font class="px12" color="black">元</font>
    </td>
    </tr>
     <tr bgcolor="white">
     <td  align="left" bgcolor="eeeeee">
     <font class="px12" color="black">&nbsp;&nbsp;费用说明：</font>
               <input type="text" name="mnynote" size="60" value=<%=objrst("mnynote")%> readonly> 
     <br><br> 
         <font class="px12" color="black">&nbsp;&nbsp;提交人：<%=name%>&nbsp;&nbsp;提交日：<%=date()%></font>
        
      </td>
    </tr>

    
    <tr align="center" bgcolor="white"> <td align="middle" bgcolor="eeeeee"><input type="submit" value="提   交" id="submit1" name="submit1"></td></tr> 
        </form>   
   
  <%end sub%>            
           
             

<!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          
q=Request.QueryString("q")

 set rs_r2=server.CreateObject("adodb.recordset")                                                                            
rs_r2.CursorLocation=2  
sql="select * from cwys_infoin where record_id='"&q&"'"
rs_r2.open sql,conn_1,3,3,1
passcode=trim(rs_r2("passcode"))
passname=trim(rs_r2("passname"))
bxcode=trim(rs_r2("bxcode"))
bxname=trim(rs_r2("bxname"))
djdate=trim(rs_r2("djdate"))
djname=trim(rs_r2("djname"))
mnydepm=trim(rs_r2("mnydepm"))
mnykm=trim(rs_r2("mnykm"))
mnynote=trim(rs_r2("mnynote"))
price=trim(rs_r2("price"))
date1=trim(rs_r2("mnytime"))
payway=trim(rs_r2("payway"))
ifhandin=trim(rs_r2("ifhandin"))
mnyyear=trim(rs_r2("mnyyear"))

fkmcode2=trim(rs_r2("mnykmcode"))

set rs_r1=server.CreateObject("adodb.recordset")                                                                            
rs_r1.CursorLocation=2                                                                            
sqlr="insert into cwys_bmglrz (passcode,passname,bxcode,bxname,djdate,djname,mnydepm,mnykm,mnynote,price,mnytime,payway,ifhandin,ifhx,mnyyear,changeid,cz,lururen,lurutime,mnykmcode) values ('"&passcode&"','"&passname&"','"&bxcode&"','"&bxname&"','"&djdate&"','"&djname&"','"&mnydepm&"','"&mnykm&"','"&mnynote&"','"&price&"','"&date1&"','"&payway&"','"&ifhandin&"','否','"&mnyyear&"','"&q&"','超标提交','"&session("emid")&"','"&date()&"','"&trim(fkmcode2)&"')"                                                                            
'Response.Write(sqlb)
rs_r1.Open sqlr,conn_1,3,3,1


fkmcode1=Request.QueryString("fkmcode1")
'Response.Write(q1)
'获得数据
passcode = trim(Request.form("hbh"))
if passcode=""then
passcode ="无"
end if

passname = trim(Request.form("hd"))
if passname=""then
passname ="无"
end if

bxcode = trim(Request.form("hbh1"))
if bxcode=""then
bxcode ="无"
end if

bxname = trim(Request.form("hd1"))
if bxname=""then
bxname ="无"
end if

djdate=date()
djname=session("emid")

mnydepm=trim(f)

mnykm=trim(Request.Form("mnykm"))

mnynote=trim(Request.Form("mnynote"))
if mnynote=""then
mnynote="无"
end if
 


payway = trim(Request.form("payway"))

ifhx="否"

mnyyear = trim(Request.form("year1"))
mnymonth = trim(Request.Form("month1"))

date1=mnyyear&"-"&mnymonth&"-"&"1"
date1=cdate(date1)

price = trim(Request.form("price"))
if price=""then
price ="0"
end if 

set rs_c=server.CreateObject("adodb.recordset")                                                                            
rs_c.CursorLocation=2                                                                            
sqlc="SELECT * FROM cwys_ed WHERE fkmcode='"&trim(fkmcode2)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'"                                                                            
rs_c.Open sqlc,conn_1,3,3,1
if rs_c.EOF then%>
<table bgcolor=cccccc width=650>

<tr bgcolor=e3e3e3>

<td><font class="px12">没有此项科目的额度,您的帐单没有入库,请联系财务为此项科目分配额度</font></td>
</tr>

</table>
<%else
isover=trim(rs_c("isover"))
'Response.Write(isover)
%>

<%

'截止用户所选月份能用的钱的总和

set rs_7=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql7="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode2)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

rs_7.Open sql7,conn_1,3,3,1
'Response.Write(sql7)
if not rs_7.EOF then


qq=cint(mnymonth)


if qq="1" then
yxmy=rs_7("jan")
date2=mnyyear&"-"&"1"&"-"&"1"
date2=cdate(date2)
end if

if qq="2" then
yxmy=rs_7("jan")+rs_7("feb")
date2=mnyyear&"-"&"2"&"-"&"1"
date2=cdate(date2)
end if

if qq="3" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")
date2=mnyyear&"-"&"3"&"-"&"1"
date2=cdate(date2)
end if

if qq="4" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")
date2=mnyyear&"-"&"4"&"-"&"1"
date2=cdate(date2)
end if

if qq="5" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")
date2=mnyyear&"-"&"5"&"-"&"1"
date2=cdate(date2)
end if

if qq="6" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")
date2=mnyyear&"-"&"6"&"-"&"1"
date2=cdate(date2)
end if


if qq="7" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")
date2=mnyyear&"-"&"7"&"-"&"1"
date2=cdate(date2)
end if

if qq="8" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")
date2=mnyyear&"-"&"8"&"-"&"1"
date2=cdate(date2)
end if

if qq="9" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")
date2=mnyyear&"-"&"9"&"-"&"1"
date2=cdate(date2)
end if

if qq="10" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")
date2=mnyyear&"-"&"10"&"-"&"1"
date2=cdate(date2)
end if

if qq="11" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")+rs_7("nov")
date2=mnyyear&"-"&"11"&"-"&"1"
date2=cdate(date2)
end if

if qq="12" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")+rs_7("nov")+rs_7("dece")
date2=mnyyear&"-"&"12"&"-"&"1"
date2=cdate(date2)
end if
'if not rs_7.EOF then
end if
%>

<%
'到所选月份用户已经使用过的钱

set rs_6=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql6="SELECT sum(price)as mnyused FROM cwys_infoin where mnykmcode='"&trim(fkmcode2)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' and mnytime<='"&date2&"' and ifhandin='是' and cz<>'删除' group by mnykmcode,mnydepm"
'Response.Write(sql6)
rs_6.Open sql6,conn_1,3,3,1
if rs_6.EOF then
mnyused=0
else

mnyused=rs_6("mnyused")
end if
%>


<%

'到用户所zai月份所在季度能用的钱的总和

set rs_77=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql77="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode2)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

rs_77.Open sql77,conn_1,3,3,1
'Response.Write(sql77)
if not rs_77.EOF then
p=cint(Month(date))

if p="1" then
yxmy1=rs_77("jan")
end if

if p="2" then
yxmy1=rs_77("jan")+rs_77("feb")
end if

if p="3" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")
end if

if p="4" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")
end if

if p="5" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")
end if

if p="6" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")
end if

if p="7" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")
end if

if p="8" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")+rs_77("aug")
end if

if p="9" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")+rs_77("aug")++rs_77("sep")
end if

if p="10" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")+rs_77("aug")++rs_77("sep")+rs_77("oct")
end if

if p="11" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")+rs_77("aug")++rs_77("sep")+rs_77("oct")+rs_77("nov")
end if

if p="12" then
yxmy1=rs_77("jan")+rs_77("feb")+rs_77("mar")+rs_77("apr")+rs_77("may")+rs_77("jun")+rs_77("jul")+rs_77("aug")++rs_77("sep")+rs_77("oct")+rs_77("nov")+rs_77("dece")
end if



end if 

%>

<%
'到所zai月份用户已经使用过的钱

set rs_66=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql66="SELECT sum(price)as mnyused FROM cwys_infoin where mnykmcode='"&trim(fkmcode2)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' and mnytime<='"&date()&"' and ifhandin='是' and cz<>'删除' group by mnykmcode,mnydepm"
'Response.Write(sql66)
rs_66.Open sql66,conn_1,3,3,1
if rs_66.EOF then
mnyused1=0
else

mnyused1=rs_66("mnyused")
end if
%>

<%
'Response.Write(isover)
if isover="False" then%>





<%
shengyu=yxmy-mnyused
chae=shengyu-price

shengyu1=yxmy1-mnyused1
chae1=shengyu1-price



%>

<%if chae>=0 and chae1>=0 then%>
<%'存
ifhandin="是"
'Response.Write("nengcun")
'Response.Write(price)%>

<%

set rs_a=server.CreateObject("adodb.recordset")                                                                            
rs_a.CursorLocation=2                                                                            
sqla="update cwys_infoin set passcode='"&passcode&"',passname='"&passname&"',bxcode='"&bxcode&"',bxname='"&bxname&"',djdate='"&djdate&"',mnydepm='"&mnydepm&"',mnykm='"&mnykm&"',mnynote='"&mnynote&"',price="&price&",mnyyear='"&mnyyear&"',mnytime='"&date1&"',payway='"&payway&"',ifhandin='"&ifhandin&"',ifhx='"&ifhx&"' where record_id='"&q&"'"

rs_a.Open sqla,conn_1,3,3,1
'Response.Write(sqla)



%>



<table>

<tr><td><font class="px12" color="red">已经提交财务,<%=mnyyear%>年度截止<%=mnymonth%>月份允许额度<%=yxmy%>元，已经使用<%=mnyused+price%>,还能使用<%=shengyu-price%>元</font><input value="确 定" type="submit" name="action" onclick="javascript:reload()"></td></tr>
</table>




 <%else%>
 <%'Response.Write("bunengcun")
ifhandin="否"
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2                                                                            
sqlb="update cwys_infoin set passcode='"&passcode&"',passname='"&passname&"',bxcode='"&bxcode&"',bxname='"&bxname&"',djdate='"&djdate&"',mnydepm='"&mnydepm&"',mnykm='"&mnykm&"',mnynote='"&mnynote&"',price="&price&",mnyyear='"&mnyyear&"',mnytime='"&date1&"',payway='"&payway&"',ifhandin='"&ifhandin&"',ifhx='"&ifhx&"' where record_id='"&q&"'"
'Response.Write(sqlb)
rs_b.Open sqlb,conn_1,3,3,1



%>


<table>
<tr><td><font class="px12" color="blue">数额超标，<font color="red">不能提交财务</font>,<%=mnyyear%>年度截止<%=mnymonth%>月份允许额度<%=yxmy%>元，已经使用<%=mnyused%>,还能使用<%=yxmy-mnyused%>元,目前需要财务调整额度<%=price-yxmy+mnyused%>元</font><input value="确 定" type="submit" name="action" onclick="javascript:reload()"></td></tr>
</table>



 <%end if%>
 
 
 
 <%else%>
  <%'存
ifhandin="是"
'Response.Write("nengcun")
'Response.Write(price)%>

<%

set rs_u=server.CreateObject("adodb.recordset")                                                                            
rs_u.CursorLocation=2                                                                            
sqlu="update cwys_infoin set passcode='"&passcode&"',passname='"&passname&"',bxcode='"&bxcode&"',bxname='"&bxname&"',djdate='"&djdate&"',mnydepm='"&mnydepm&"',mnykm='"&mnykm&"',mnynote='"&mnynote&"',price="&price&",mnyyear='"&mnyyear&"',mnytime='"&date1&"',payway='"&payway&"',ifhandin='"&ifhandin&"',ifhx='"&ifhx&"',isover='是' where record_id='"&q&"'"
'Response.Write(sqlu)
rs_u.Open sqlu,conn_1,3,3,1


%>
<%

set rs_o=server.CreateObject("adodb.recordset")                                                                            
rs_o.CursorLocation=2                                                                            
sqlo="update cwys_ed set isover='0' where depar='"& f &"' and ys_year='"&mnyyear&"' and fkmcode='"&trim(fkmcode2)&"'"                                                                            
'Response.Write(sqla)
rs_o.Open sqlo,conn_1,3,3,1


%>
<table>

<tr><td><font class="px12" color="red">已经提交财务,此次提交成功是由于财务部允许此项费用超标,<%=mnyyear%>年度截止<%=mnymonth%>月份允许额度<%=yxmy%>元，已经使用<%=mnyused+price%>,还能使用0元<input value="确 定" type="submit" name="action" onclick="javascript:reload()"></font></td></tr>
</table>
<%end if%>
<%end if%> 
<%end sub
%>






<%'主过程
                                                                                                                        
Select case Request.QueryString("todo")                                                                         
case""
list()
         
       case"02" 
       save_data()  
       list()                                               
End Select                                                                       
  %>
                  </tbody>
                </table>
                </td>
                </tr>
                </tbody>
                </table>
   <br>
    

      
    </td>
  </tr>
</table>
</body>
</html>