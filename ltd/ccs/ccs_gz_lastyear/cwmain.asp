<%@ Language=VBScript %>

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
<title>深圳公司预算管理系统</title>
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
function edit(sn)
{
tzurl="hx_edit.asp?sn="+sn;
window.open(tzurl,"xx","width=600,left=200,top=10,height=300,scrollbars=no");
}
function hx(sn)
{
tzurl="hx_hx.asp?sn="+sn;
window.open(tzurl,"xx","width=600,left=200,top=10,height=300,scrollbars=no");
}
function back(sn)
{
tzurl="hx_back.asp?sn="+sn;
window.open(tzurl,"xx","width=600,left=200,top=10,height=300,scrollbars=no");
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
function xx(s,y,x,a,b,c) { //v3.0
   //s=depar.value ;
   //y=nian.value ;
   surl="cwmain.asp?depar="+s+"&passcode="+y+"&bxcode="+x+"&payway="+a+"&price="+b+"&tabid="+c;
   window.location.href (surl);
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
                <td><a href="cwmain.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/baoxiao1.gif',1)"><img src="images/baoxiao.gif" width="120" height="24" border="0" name="images1"></a></td>
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
                <td><a href><img src="images/bxsp.gif" width="138" height="28" border="0"></a></td>
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
                     <td><a href="cwmain.asp"><font class="px14" color="#FFFFFF">未核销查询</font></a></td>
                      </tr>
                      
                      <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="query1.asp"><font class="px14" color="#FFFFFF">已核销查询</font></a></td>
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
  
     <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="30" cellSpacing="0" cellPadding="0" width="850" border="0">


  
     <tr height="30" bgcolor="#E3E9EE" align="left">
     <td >    <font class="px12" color="black">&nbsp;&nbsp;部门：<select name="depar" style="HEIGHT: 22px; WIDTH: 100px"> 
    <option selected value="全部">全部</option>
    <%
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="Select * From CompanyLocale where companyid not in (6,7,11,17,21,26,27,28) order by companyname"
   obj.Source=sql
   obj.Open
    obj.MoveFirst 
    while not obj.EOF
    %>
    <option value="<%=obj("CompanyName")%>"><%=obj("CompanyName")%></option>
    <%
    obj.movenext
    wend
    obj.close
    %>	
    </select>&nbsp;&nbsp;
     经办人员工号：<INPUT  type="text"  name="passcode" size=8>
     &nbsp;&nbsp;
     报销人员工号：<INPUT  type="text"  name="bxcode" size=8>
     &nbsp;&nbsp;</br>
     &nbsp;&nbsp;付款方式：<select name="payway">
         <OPTION  value="现金"><font color=black class="px14">现金</OPTION>
         <OPTION  value="银行"><font color=black class="px14">银行</OPTION>
         <OPTION  value="内部往来"><font color=black class="px14">内部往来</OPTION>
         <OPTION  value="转帐"><font color=black class="px14">转帐</OPTION>
         </select>
     &nbsp;&nbsp;金额：<INPUT  type="text"  name="price" size=8>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp
     报销单号：<INPUT  type="text"  name="tabid" size=12>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     <input type="button" name="action" value="查询" onclick="javascript:xx(depar.value,passcode.value,bxcode.value,payway.value,price.value,tabid.value)">
     </td>
     </tr>             
    </table>
    
      <%
      depar=trim(Request.QueryString("depar"))
      if depar<>"全部" then
      sql1="and mnydepm='"&depar&"'"
      end if
      
      passcode=trim(Request.QueryString("passcode"))
      if passcode<>"" then
      sql2="and passcode='"&passcode&"'"
      end if
      
      bxcode=trim(Request.QueryString("bxcode"))
      if bxcode<>"" then
      sql3="and bxcode='"&bxcode&"'"
      end if
      
      payway=trim(Request.QueryString("payway"))
      if payway<>"全部" then
      sql4="and payway='"&payway&"'"
      end if
      
      price=trim(Request.QueryString("price"))
      if price<>"" then
      sql5="and price=convert(money,'"&price&"')"
      end if
      
      tabid=trim(Request.QueryString("tabid"))
      if tabid<>"" then
      sql6="and tabid='"&tabid&"'"
      end if      
      %> 
    <%
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="select * from cwys_infoin where ifhandin='是' and ifhx='否'"+sql1+sql2+sql3+sql4+sql5+sql6+"order by djdate desc"
   obj.Source=sql
   obj.Open
    %> 
   <% if not obj.eof then%>  
   

     <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 0px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid"  cellSpacing="1" cellPadding="1" width="850" border="0">
     <tr height="20" bgcolor="#E3E9EE" align="left">
     
     <td ><font class="px12" color="blue">操作</font></td>
     <td ><font class="px12" color="black">部门</font></td>
     <td ><font class="px12" color="black">费用期间</font></td>
     <td ><font class="px12" color="black">付款方式</font></td>
     <td ><font class="px12" color="black">科目</font></td>
     <td ><font class="px12" color="black">金额</font></td>
     <td ><font class="px12" color="black">简要说明</font></td>
     <td ><font class="px12" color="black">报销人</font></td>
     <td ><font class="px12" color="black">登记人</font></td>
     <td ><font class="px12" color="black">经办人</font></td>
     <td ><font class="px12" color="black">登记时间</font></td>
     <td ><font class="px12" color="blue">核销</font></td>
     <td ><font class="px12" color="blue">退回</font></td>
     </tr>   
   <%
   obj.MoveFirst 
   while not obj.EOF
   %>  
   
   <%
	'Set Conn = Server.CreateObject("ADODB.Connection")
	'Conn.Open Application("OledbStr") 
	'Set obj1=server.CreateObject ("ADODB.Recordset")
	'obj1.LockType=3
	'obj1.CursorType=3
	'set obj1.activeConnection=Conn 
	'obj1.Source="select * FROM cwys_km where fkmshuom='"&obj("mnykm")&"' and depar='"&obj("mnydepm")&"' and nian='"&obj("mnyyear")&"'"
	'obj1.Open 
   %>
   <% 'if obj1.eof then%> 
   <%'else%>
     <tr height="20" bgcolor="#E3E9EE" align="left">
     <td ><input type="button"  value="修改" onclick="javascript: edit(<%=obj("record_id")%>)" id=button1 name=button1></td>
     <td ><font class="px12" color="black"><%=obj("mnydepm")%></font></td>
     <td ><font class="px12" color="black"><%=obj("mnyyear")%>年<%=month(obj("mnytime"))%>月</font></td>
     <td ><font class="px12" color="black"><%=obj("payway")%></font></td>
     <td ><font class="px12" color="black"><%=obj("mnykm")%></font></td>
     <td ><font class="px12" color="black"><%=obj("price")%></font></td>
     <td ><font class="px12" color="black"><%=obj("mnynote")%></font></td>
     <td ><font class="px12" color="black"><%=obj("bxname")%></font></td>
     <td ><font class="px12" color="black"><%=obj("djname")%></font></td>
     <td ><font class="px12" color="black"><%=obj("passname")%></font></td>
     <td ><font class="px12" color="black"><%=obj("djdate")%></font></td>
     <td ><input type="button" value="核销" onclick="javascript: hx(<%=obj("record_id")%>)" id=button1 name=button1></td>
     <td ><input type="button" style="background-color: #E0E0E0; color: red;class=px14; border: 0 solid #9E9E9E" value="退回" onclick="javascript: back(<%=obj("record_id")%>)" id=button1 name=button1></td>

     </tr>
    <%'end if%>  
    <%'obj1.Close%> 
    
   <%
   obj.movenext
   wend
   %>	     
    </table>
   <%end if%> 
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


