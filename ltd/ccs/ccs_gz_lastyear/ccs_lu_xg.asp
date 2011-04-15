
<%@ Language=VBScript %>


  
<html>
<head>
<title>深圳公司预算管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<%q=Request.QueryString("q")
%>
<% 
 

e=trim(session("emid"))
t=trim(session("loginid"))
Response.Write(e)
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
sql3="select * from companylocale where companyid='"&cid&"'"
rs_3.Open sql3,conn_1,3,3,1

f=trim(rs_3("companyname"))


Response.Write(f)
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/obj_hed1.gif">
      <table border="0" cellspacing="0" cellpadding="0" width="800">
        <tr>
          <td width="10"><img src="images/spacer.gif" width="10" height="35"></td>
          <td width="171"><img src="images/obj_maintitle.gif" width="171" height="30" border="0"></td>
          <td width="10"><img src="images/spacer.gif" width="10" height="35"></td>
          <td width="605"> 
            <table border="0" cellspacing="0" cellpadding="3" name="menubutton">
              <tr>
                <td>&nbsp;</td>
                 <td>&nbsp;</td>

                <td><a href="ccs_ysqk_index.asp?" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/bt_01_off.gif',1)"><img src="images/bt_01_on.gif" width="120" height="25" border="0" name="images1"></a></td>
              
                <td><a href="ccs_input_index.asp?" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/bt_02_off.gif',1)"><img src="images/bt_02_on.gif" width="120" height="25" border="0" name="images2"></a></td>
                <td><a href="ccs_yskh_index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images3','','images/bt_03_off.gif',1)"><img src="images/bt_03_on.gif" width="120" height="24" border="0" name="images3"></a></td>
                <td><a href="ccs_xtwf.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/bt_04_off.gif',1)"><img src="images/bt_04_on.gif" width="120" height="24" border="0" name="images4"></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td bgcolor="#FF6600"><img src="images/spacer.gif" width="10" height="2"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="700">
  <tr> 
    <td width="140"><img src="images/obj_hed2_left.jpg" width="140" height="98"></td>
    <td width="30"><img src="images/obj_hed2_center2.jpg" width="30" height="98"></td>
    <td width="470" valign="top" background="images/obj_hed2_right.jpg"> 
      <table width="300" border="0" cellspacing="0" cellpadding="0" background>
        <tr>
          <td><img src="images/spacer.gif" width="330" height="5"></td>
          <td><img src="images/spacer.gif" width="130" height="5"></td>
        </tr>
        <tr>
          <td valign="top">
            <table border="0" cellspacing="0" cellpadding="0" name="banner">
              <tr> 
                <td colspan="2"><img src="images/spacer.gif" width="10" height="24"></td>
              </tr>
              <tr> 
                <td><img src="images/ba1_point.gif" width="35" height="39"></td>
                <td><img src="images/lizztp1.gif" width="205" height="39"></td>
              </tr>
            </table>
          </td>
          <td align="right">
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/spacer.gif" width="20" height="53"></td>
              </tr>
              <tr>
                <td><a href><img src="images/lizztp2.gif" width="138" height="28" border="0"></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
    </td>
    <td width="470" valign="top"><img src="images/obj_hed2_right-.jpg" width="60" height="73"> 
    </td>
  </tr>
</table>
<table width="740" border="0" cellspacing="0" cellpadding="0" height="90%">
  <tr>
    <td bgcolor="#000033" width="140" valign="top"> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/obj_left.jpg"> 
            <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="images/spacer.gif" width="140" height="45"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="5" cellpadding="0" width="140">
        <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
      
        <tr> 
          <td> 
            <table width="130" border="0" cellspacing="0" cellpadding="1">
            
              <%objrst.Source ="select fkmshuom from cwys_km where depar = '"&f&"' and nian='"&year(date())&"' order by sn" 
                       objrst.Open 
                       
                     
                      while not objrst.EOF %>
                     <tr> 
                      <td width="6" valign="top"><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                     <td width="116"><a href="ccs_ser.asp?km=<%=objrst("fkmshuom")%>"><font class="px14" color="#FFFFFF"><%=objrst("fkmshuom")%></font></a></td>
                     </tr>
                     <%objrst.MoveNext %> 
                       <% wend %>
                       <%objrst.Close%>
                     
            
            
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
       
       
      </table>
    </td>
    <td width="5"><img src="images/spacer.gif" width="10" height="5"></td>
    <td width="595" valign="top"> 
    <form method="post" action="ccs_lu_xg.asp?todo=02&q=<%=q%>" id=form1 name=form1>   
     <table style="BORDER-RIGHT: #4e4c71 1px solid; BORDER-TOP: #4e4c71 1px solid; BORDER-LEFT: #4e4c71 1px solid; BORDER-BOTTOM: #4e4c71 1px solid" height="92%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4e4c71" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff" class="px14">&nbsp;科目录入</font> </td></tr>
        <tr>
          <td vAlign="top" width="100%">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
              
              <tr bgColor="#d8d9de" height="20">
               <%objrst.Source ="select * from cwys_infoin where record_id = '"&q&"'" 
                       objrst.Open
                       'Response.Write(objrst.Source)
                      %>
    <td align="right" >&nbsp;</td></tr>
   <tr>
    <td align="center" bgcolor="#d8d9de"><font class="px12" color="black">您目前进入的是<font color=red><%=q%></font>号纪录<font color=red><%=objrst("mnykm")%></font>修改界面</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>

                       
<tr width=100%>
     <td  align="left">
     <font class="px12" color="black">&nbsp;&nbsp;经办人员工号：</font>
     <input id="hbh" type="text" name="hbh" size="5" onchange="javascript:disname()" value=<%=objrst("passcode")%>>    
    
     <font class="px12" color="black">&nbsp;&nbsp;经办人姓名：</font>
     <input id="hd" type="text" name="hd" size="10" value=<%=objrst("passname")%>>
     
     <font class="px12" color="black">&nbsp;&nbsp;费用控制部门：</font>
     <input type="text" name="mnydepm" size="10" readonly value=<%=f%>>
<br>
 
     <font class="px12" color="black">&nbsp;&nbsp;报销人员工号：</font>
     <input type="text" name="hbh1" size="5" onchange="javascript:disname1()" value=<%=objrst("bxcode")%>> 
     
     <font class="px12" color="black">&nbsp;&nbsp;报销人姓名：</font>
     <input type="text" name="hd1" size="10" value=<%=objrst("bxname")%>>
     
     <font class="px12" color="black">&nbsp;&nbsp;费用科目：</font>
     <input type="text" name="mnykm" size="10" readonly value=<%=objrst("mnykm")%>>
<%set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")                                                                           
set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2                                                                            
sql2="SELECT fkmcode FROM cwys_km WHERE fkmshuom='"& objrst("mnykm") &"' and depar='"&f&"' and nian='"&year(date())&"'"                                                                            
rs_2.Open sql2,conn_1,3,3,1
if not rs_2.EOF then
fkmcode1=rs_2("fkmcode")
'Response.Write(fkmcode1)
else
fkmcode1="无"
end if
%>
<font class="px12"><%=fkmcode1%></font>      
<br>
   
      </td>
    </tr>


    <tr>
      <td align="left">
    <font class="px12" color="black">&nbsp;&nbsp;费用期间：</font>
        <select size="1" name="year1"><option selected><%=objrst("mnyyear")%>
         <OPTION  value="<%=cstr(year(date)+1)%>"><font color=black class="px14"><%=cstr(year(date)+1)%></OPTION>
             <OPTION selected value="<%=cstr(year(date))%>"><font color=black class="px14"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><font color=black class="px14"><%=cstr(year(date)-1)%></OPTION>
                       <OPTION  value="<%=cstr(year(date)-2)%>"><font color=black class="px14"><%=cstr(year(date)-2)%></OPTION>

            </select>
            <font color=black class="px14">年</font>
        <select name="month1" style="HEIGHT: 22px; WIDTH: 43px"><option selected> <%=month(objrst("mnytime"))%>
        <%dd=month(date)+1
        
          cha=dd-1
          while cha>0 
    %>
        <option><%=cha%></option>
        
        <%
        cha=cha-1
        wend%></select>
        
      <font class="px12" color="black">月</font>
     
     <font class="px12" color="black">&nbsp;&nbsp;付款方式：</font>
        <select size="1" name="payway"><option selected><%=objrst("payway")%>
         <OPTION  value="现金"><font color=black class="px14">现金</OPTION>
         <OPTION  value="银行"><font color=black class="px14">银行</OPTION>
         <OPTION  value="内部往来"><font color=black class="px14">内部往来</OPTION>
         <OPTION  value="预提"><font color=black class="px14">预提</OPTION>

            </select>
                
     <font class="px12" color="black"><br>&nbsp;&nbsp;金额：</font>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="price" size="10" value=<%=objrst("price")%>><font class="px12" color="black">元</font>
    </td>
    </tr>
     <tr>
     <td  align="left">
     <font class="px12" color="black">&nbsp;&nbsp;费用说明：</font>
               <input type="text" name="mnynote" size="80" value=<%=objrst("mnynote")%>> 
     <br><br> 
         <font class="px12" color="black">&nbsp;&nbsp;修改人：<%=name%>&nbsp;&nbsp;修改日：<%=date()%></font>
        
      </td>
    </tr>

    <tr>
    <td align="right">&nbsp;</td></tr>
    <tr align="center"> <td align="middle" ><input type="submit" value="提   交" id="submit1" name="submit1"></td></tr> 
        </form>   
    <tr height=5>
    <td align="right" height=5 ><a href="ccs_ser.asp?km=<%=objrst("mnykm")%>"><font class="px12" color="black">返回修改删除列表</a>&nbsp;&nbsp;&nbsp;</td></tr>
              
           
             

<!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          
q=Request.QueryString("q")
Response.Write(q1)
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

price = trim(Request.form("price"))
if price=""then
price ="0"
end if 


%>



<%





'到用户所选月份能用的钱的总和 
set rs_7=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql7="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode1)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

rs_7.Open sql7,conn_1,3,3,1
'Response.Write(sql7)
if not rs_7.EOF then

i=5
g=1
v=cint(mnymonth)
while g<=v
yxmy=rs_7(i)+yxmy
i=i+1
g=g+1
wend

else
yxmy=0
end if
Response.Write(yxmy)

%>



<%
'到所选月份用户已经使用过的钱

set rs_6=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql6="SELECT mnykm,mnydepm,sum(price)as mnyused FROM cwys_infoin where mnykm='"&trim(fkmcode1)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' group by mnykm,mnydepm"

rs_6.Open sql6,conn_1,3,3,1
if rs_6.EOF then
mnyused=0
else
'Response.Write(sql6)
mnyused=rs_6("mnyused")
end if
%>
<%
shengyu=yxmy-mnyused
chae=shengyu-price
if chae<0 then
%>
<table>
<br>
<tr><td><font class="px12">已经修改，但数额超标，不能提交财务核销</font></td></tr>
</table>
<%'Response.Write("bunengcun")
ifhandin="否"
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2                                                                            
sqlb="update cwys_infoin set passcode='"&passcode&"',passname='"&passname&"',bxcode='"&bxcode&"',bxname='"&bxname&"',djdate='"&djdate&"',mnydepm='"&mnydepm&"',mnykm='"&mnykm&"',mnynote='"&mnynote&"',price='"&price&"',mnyyear='"&mnyyear&"',mnytime='"&date1&"',payway='"&payway&"',ifhandin='"&ifhandin&"',ifhx='"&ifhx&"' where record_id='"&q&"'"
'Response.Write(sqlb)
rs_b.Open sqlb,conn_1,3,3,1


%>

 <%else%>
 <%'存
ifhandin="是"
'Response.Write("nengcun")
'Response.Write(price)%>
<br><br>
<table>

<tr><td><font class="px12">已经修改，已经提交财务核销</font></td></tr>
</table>
<%

set rs_a=server.CreateObject("adodb.recordset")                                                                            
rs_a.CursorLocation=2                                                                            
sqla="update cwys_infoin set passcode='"&passcode&"',passname='"&passname&"',bxcode='"&bxcode&"',bxname='"&bxname&"',djdate='"&djdate&"',mnydepm='"&mnydepm&"',mnykm='"&mnykm&"',mnynote='"&mnynote&"',price='"&price&"',mnyyear='"&mnyyear&"',mnytime='"&date1&"',payway='"&payway&"',ifhandin='"&ifhandin&"',ifhx='"&ifhx&"' where record_id='"&q&"'"

rs_a.Open sqla,conn_1,3,3,1
'Response.Write(sqla)

%>

 <%end if%>
<%end sub
%>






<%'主过程
                                                                                                                        
Select case Request.QueryString("todo")                                                                         

         
       case"02" 
       save_data()  
                                                      
End Select                                                                       
  %>
                  </tbody>
                </table>
                </td>
                </tr>
                </tbody>
                </table>
   <br>
    
      <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><font class="px12">Copyright 2006, 中国南方航空深圳公司信息工程部</font></td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
</body>
</html>
