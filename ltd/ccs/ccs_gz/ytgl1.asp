<%@ Language=VBScript %>

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

function xx(s,y) { 
   surl="ysgl.asp?depar="+s+"&nian="+y;
   window.location.href (surl);
}
function edit(sn,nian)
{
tzurl="ys_edit.asp?sn="+sn+"&nian="+nian;
window.open(tzurl,"xx","width=600,left=200,top=10,height=400,scrollbars=no");
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
                <td><a href="ytgl.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/yuti1.gif',1)"><img src="images/yuti.gif" width="120" height="24" border="0" name="images1"></a></td>
                <td><a href="cx/ccs_gscxy_index.asp" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/gltj2.jpg',1)"><img src="images/gltj1.jpg" width="120" height="24" border="0" name="images4"></a></td>
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
                <td><a href><img src="images/ysgl.gif" width="138" height="28" border="0"></a></td>
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
            
                       <tr>
                     <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                     <td><a href="ytgl.asp"><font class="px14" color="#FFFFFF">预提</font></a></td>
                      </tr>
                      <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="yt_ch.asp"><font class="px14" color="#FFFFFF">冲销</font></a></td>
                       </tr>
                      <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="yt_ts.asp"><font class="px14" color="#FFFFFF">托收</font></a></td>
                       </tr>
                      <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="yt_cx.asp"><font class="px14" color="#FFFFFF">查询</font></a></td>
                      </tr>
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
<table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="60" cellSpacing="1" cellPadding="1" width="550" border="1">
<%
	yue=Request.form("yue")
	nian=Request.form("nian")
	depar=trim(Request.form("depar"))
	kemu=trim(Request.form("kemu"))
	edu=Request.form("edu")
	note=trim(Request.form("note"))
	yuefen=cdate(nian+"-"+yue+"-1")
	jd1time=cdate(nian+"-3-1")

	jd2time=cdate(nian+"-6-1")

	jd3time=cdate(nian+"-9-1")

	jd4time=cdate(nian+"-12-1")
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr")
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select distinct * from cwys_km where  fkmcode='"&kemu&"' and nian='"&nian&"' "
   obj.Source=sql
   obj.Open
   %>
   <% 
   if not obj.eof then
   fkmshuom=trim(obj("fkmshuom"))
   else
   fkmshuom=""
   end if
   obj.Close
%>

<% if depar=""  then %> 
<tr align="center" >
<td align="center" >
	<font class=px14 color=blue>部门为空，请选择部门。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
<%elseif  kemu="" then%>
<tr align="center" >
<td align="center" >
	<font class=px14 color=blue>预算科目为空，请选择预算科目。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
<%elseif edu="" then%>
<tr align="center" >
<td align="center" >
	<font class=px14 color=blue>预算指标为空，请输入预算指标。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
<%elseif note="" then%>
<tr align="center" >
<td align="center" >
	<font class=px14 color=blue>摘要为空，请输入摘要内容。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
<%else%>

<%'在这里添加判断，判断预提金额是否大于该月额度的剩余值。
    if yue>=1 and yue<=3 then
	kytime=cdate(nian+"-3-1")
	jd=1
	elseif yue>=4 and yue<=6 then
	kytime=cdate(nian+"-6-1")
	jd=2
	elseif yue>=7 and yue<=9 then
	kytime=cdate(nian+"-9-1")
	jd=3
	else
	kytime=cdate(nian+"-12-1")
	jd=4
	end if
	
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr")
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select ISNULL(SUM(price), 0)  from cwys_infoin where  mnykmcode='"&kemu&"' and mnytime<='"&kytime&"' and mnyyear='"&nian&"' and mnydepm='"&depar&"' and ifhandin='是' and  cz<>'删除'"
   obj.Source=sql
   obj.Open
   mnysum=obj(0)
   obj.Close
   
   Set obj11=server.CreateObject ("ADODB.Recordset")
	obj11.LockType=3
	obj11.CursorType=3
	set obj11.activeConnection=objConn
	sql11="select ISNULL(SUM(price), 0) FROM cwys_infoin where mnydepm='"&depar&"' and mnytime<='"&jd1time&"' and mnyyear='"&nian&"'and mnykmcode='"&kemu&"' and  cz<>'删除' and ifhandin='是' "
	obj11.Source=sql11
	obj11.Open
	jd1=obj11(0)
	obj11.Close

	Set obj12=server.CreateObject ("ADODB.Recordset")
	obj12.LockType=3
	obj12.CursorType=3
	set obj12.activeConnection=objConn
	sql12="select ISNULL(SUM(price), 0) FROM cwys_infoin where mnydepm='"&depar&"' and mnytime<='"&jd2time&"' and mnyyear='"&nian&"' and mnykmcode='"&kemu&"' and  cz<>'删除' and ifhandin='是'  "
	obj12.Source=sql12
	obj12.Open
	jd2=obj12(0)
	obj12.Close

	Set obj13=server.CreateObject ("ADODB.Recordset")
	obj13.LockType=3
	obj13.CursorType=3
	set obj13.activeConnection=objConn
	sql13="select ISNULL(SUM(price), 0) FROM cwys_infoin where mnydepm='"&depar&"' and mnytime<='"&jd3time&"' and mnyyear='"&nian&"' and mnykmcode='"&kemu&"' and  cz<>'删除' and ifhandin='是' "
	obj13.Source=sql13
	obj13.Open
	jd3=obj13(0)
	obj13.Close

	Set obj14=server.CreateObject ("ADODB.Recordset")
	obj14.LockType=3
	obj14.CursorType=3
	set obj14.activeConnection=objConn
	sql14="select ISNULL(SUM(price), 0) FROM cwys_infoin where mnydepm='"&depar&"' and mnytime<='"&jd4time&"' and mnyyear='"&nian&"' and mnykmcode='"&kemu&"' and  cz<>'删除' and ifhandin='是' "
	obj14.Source=sql14
	obj14.Open
	jd4=obj14(0)
	obj14.Close
   
   
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr")
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select * from cwys_ed where fkmcode='"&kemu&"' and ys_year='"&nian&"' and depar='"& depar &"'"
   obj.Source=sql
   obj.Open
   if obj.eof then
   %>
   <tr align="center" >
<td align="center" >
	<font class=px14 color=blue>还没有给<%=depar%>的<%=fkmshuom%>科目分配额度。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>
  
<% 
   else  
%>

<%
jd1ky=obj("jan")+obj("feb")+obj("mar")
jd2ky=obj("jan")+obj("feb")+obj("mar")+obj("apr")+obj("may")+obj("jun")
jd3ky=obj("jan")+obj("feb")+obj("mar")+obj("apr")+obj("may")+obj("jun")+obj("jul")+obj("aug")+obj("sep")
jd4ky=obj("niandu")
if yue>=1 and yue<=3 then
keyong=obj("jan")+obj("feb")+obj("mar")
elseif yue>=4 and yue<=6 then
keyong=obj("jan")+obj("feb")+obj("mar")+obj("apr")+obj("may")+obj("jun")
elseif yue>=7 and yue<=9 then
keyong=obj("jan")+obj("feb")+obj("mar")+obj("apr")+obj("may")+obj("jun")+obj("jul")+obj("aug")+obj("sep")
else
keyong=obj("niandu")
end if
obj.Close
%>
<%
   shengyu=keyong-mnysum
if jd=1 then
jd1sy=jd1ky-jd1-edu
jd2sy=jd1ky-jd2-edu
jd3sy=jd1ky-jd3-edu
jd4sy=jd1ky-jd4-edu
end if
if jd=2 then
jd2sy=jd2ky-jd2-edu
jd3sy=jd3ky-jd3-edu
jd4sy=jd4ky-jd4-edu
end if
if jd=3 then
jd3sy=jd3ky-jd3-edu
jd4sy=jd4ky-jd4-edu
end if
if jd=4 then
jd4sy=jd4ky-jd4-edu
end if
   
   
%>


<%'将操作写入日志


descr="预提"+depar+cstr(month(mnytime))+"月份"+fkmshuom+price+"元。"
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2  
sqlb="insert into cwys_log (operation,descr,type,operatetime,operator) values ('预提','"&descr&"','预提','"&now&"','"&session("emid")&"')" 
rs_b.Open sqlb,objConn,3,3,1 
'rs_b.Close

%>
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
obj1("mnykm")=fkmshuom
obj1("mnynote")=note
obj1("price")=edu
obj1("payway")="预提"
obj1("ifhx")="是"
obj1("ifhandin")="是"
obj1("mnytime")=cdate(nian+"-"+yue+"-1")
obj1("mnyyear")=nian
obj1("cz")="录入"
                             
obj1.Update                                                                          
objConn1.CommitTrans 
obj1.close  
%>
<tr align="center" >
<td align="center" >
	<font class=px14 color=blue>预提成功。截至<%=nian%>年<%=jd%>季度的预算额度剩余<%=shengyu-edu%>元。</font></p>
	<input type="button" name="button" value="返回" onclick="JavaScript:history.go(-1)">
</td>
</tr>


<%
end if
end if
%>
      
</table>
   <br>
    <br>
     <br>
      <br>
       <br>
        <br>
         <br>
          <br>
           <br>
            <br>
             <br>



              <br> <br> <br>
               <br> <br> <br>
      <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><font class="px12">Copyright 2006, 中国南方航空深圳公司信息工程部</font></td>
        </tr>
      </table>
      <p>　</p>
    </td>
  </tr>
</table>
</body>
</html>


