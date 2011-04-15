<%@ Language=VBScript %>

<% 
 
'if trim(session("UID"))<>"" then
 '  dim objD
 '  set ObjD=server.CreateObject ("Com_UserManage1.ClsUserManage1")
 '      VerifyOk=objD.VerifyUserFunction (session("UID"),"xxys")
 '  if VerifyOk=false then
 '     session("errorNo")="000002"
 '     Response.Redirect "../../sorry/sorry.asp"
 '  end if   
 'else
 '  session("errorNo")="000001"
 '  Response.Redirect "../../sorry/sorry.asp"
'end if 
 
  'dep=Request.QueryString ("dep")

  
  year1=year(date)
  date1=cdate(year1 & "-" & "1" & "-"&"1")
  date2=cdate(year1 & "-" & "3" & "-"&"1")
  date3=cdate(year1 & "-" & "6" & "-"&"1")
  date4=cdate(year1 & "-" & "9" & "-"&"1")
  date5=cdate(year1 & "-" & "12" & "-"&"1")
  riqi=cdate(year1 & "-" & "1" & "-"&"1")
  riqi1=cdate(year1 & "-" & "12" & "-"&"1")
  km=Request.QueryString ("km")
  session("km")=km
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
                  <%objrst.Source ="select distinct kem,sx from shenzhencwys_dep where dep='"& session("dep") &"' order by sx" 
                    objrst.Open 
                    objrst.MoveFirst
                    k=objrst("kem") 
                    objrst.Close      
                       %>
                <td><a href="ccs_ysqk_index.asp?km=<%=k%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/bt_01_off.gif',1)"><img src="images/bt_01_on.gif" width="120" height="25" border="0" name="images1"></a></td>
                  <%objrst.Source ="select distinct kem,sx from shenzhencwys_dep where dep='"& session("dep") &"' order by sx" 
                    objrst.Open 
                    objrst.MoveFirst
                    k=objrst("kem") 
                    objrst.Close      
                       %>
                <td><a href="ccs_input_index.asp?km=<%=k%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/bt_02_off.gif',1)"><img src="images/bt_02_on.gif" width="120" height="25" border="0" name="images2"></a></td>
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
                <td><a href><img src="images/lizztp5.gif" width="138" height="28" border="0"></a></td>
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
            
              <%objrst.Source ="select distinct kem,sx from shenzhencwys_dep where dep='"& session("dep") &"' order by sx" 
                       objrst.Open 
                       objrst.MoveFirst
                     
                      while not objrst.EOF %>
                     <tr> 
                      <td width="6" valign="top"><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                     <td width="116"><a href="ccs_ysqk_index.asp?km=<%=objrst("kem")%>"><font class="px14" color="#FFFFFF"><%=objrst("kem")%></font></a></td>
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
   <form method="post" action="ccs_ysqk_cx.asp" id="form1" name="form1">   
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="20" cellSpacing="0" cellPadding="0" width="775"  border="0">
  <tbody>
      
   <tr>
    <td align="right" >&nbsp;</td></tr>
   <tr>
    <td align="center" ><font class="px12" color="black">您目前进入的是<%=year(date)%>年<font color=red><%=km%></font>科目预算完成情况页面</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>
     <tr>
    <td align="right" >
    
    <table>
      <tr><td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="px12" color="black">  <font color="red">请选择财务预算年份&nbsp;<select name="year1" style="HEIGHT: 22px; WIDTH: 57px"> <option selected><%=year(date)%></option>
     <option>2007</option><option>2008</option><option>2009</option><option>2010</option></select>
      <font color="#ff0000"></font>
      </td>
          <td align="right"><input type="submit" value="GO" id="submit1" name="submit1"></td></tr>    
    </table>
    </td></tr>
   </tbody></table>
   <% objrst.Source = "select * from shenzhencwys  where  dep='"& session("dep") &"' and kem='" & km &"' and bztime>='"& riqi &"' and bztime<='"& riqi1 &"' order by fkem,bztime desc"
   objRst.Open 
   if objrst.EOF and objrst.BOF   then %>    
     <P> 
     <P align=center><font class="px12" color="black"><FONT color=crimson><STRONG><%=km%>科目没有报帐信息
   <%else%>
   
<table align="center"  cellSpacing="0" cellPadding="0" width="777" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="73%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff"><b></b></font> <font color=white><a href="ccs_ysqk_mon.asp?km=<%=km%>&month=1"><font class="px12" color=white>1月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=2"><font class="px12" color=white>2月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=3"><font class="px12" color=white>3月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=4"><font class="px12" color=white>4月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=5"><font class="px12" color=white>5月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=6"><font class="px12" color=white>6月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=7"><font class="px12" color=white>7月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=8"><font class="px12" color=white>8月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=9"><font class="px12" color=white>9月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=10"><font class="px12" color=white>10月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=11"><font class="px12" color=white>11月</a> <a href="ccs_ysqk_mon.asp?km=<%=km%>&month=12"><font class="px12" color=white>12月</a></font></td></tr>
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#ecf7fd" height="20">
                <td align="middle"  ><font class="px12" color="black">帐目时间</td>
                <td align="middle" ><font class="px12" color="black">科目</td>
                <td align="middle"  ><font class="px12" color="black">子科目</td>
                <td align="middle"  ><font class="px12" color="black">金额</td>
                <td align="middle"  ><font class="px12" color="black">简要说明</td>
                <td align="middle"  ><font class="px12" color="black">录入人</td>
             
                </tr>
                <%
                i=1
                do while not objrst.EOF and i<=30 %>
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year(objrst("bztime"))%>年<%=month(objrst("bztime"))%>月</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("kem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkem")%></td>
                <td align="middle"  ><font class="px12" color="black"><%=objrst("ysmoney")%></td>
                <td align="middle"  ><font class="px12" color="black"><%=objrst("meno")%></td>
                <td align="middle"  ><font class="px12" color="black"><%=objrst("name")%></td>
                </tr>
                <%i=i+1
                objrst.MoveNext 
                loop%> 
          <%objrst.Close %>
             <tr bgColor="#ecf7fd" height="20" >
                <td align="middle"  colspan="6"><font class="px12" color="black"><font color=red><%=km%>科目详细统计</font></td>
               
            </tr>
            <tr bgColor="#ecf7fd" height="20">
                <td align="middle"  ><font class="px12" color="black">时间</td>
                <td align="middle" ><font class="px12" color="black">科目</td>
                <td align="middle"  ><font class="px12" color="black">金额</td>
                <td align="middle"  ><font class="px12" color="black">上级核定数</td>
                <td align="middle"  colspan="2"><font class="px12" color="black">完成比例</td>
             </tr>
             <% 
               objrst.Source = "select * from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>='"& riqi &"' and yea<='"& riqi1 &"'"
               objRst.Open
               if not(objrst.EOF and objrst.BOF) then 
                objrst.Close  '按科目方式进行上级核定数考核            
             %>
            <% objrst.Source = "select kem,sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and kem='"& km &"' and bztime>='"& riqi &"' and bztime<='"& riqi1 &"' group by kem"
               objRst.Open %>
         
            <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年总计</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
             <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>='"& riqi &"' and yea<='"& riqi1 &"' group by kem"
                objRst1.Open
                if not objrst1.BOF   then 
                 if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
                <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                 <%else%>
                 <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2">&nbsp;</td> 
                <%end if%> 
                <%else%>
                 <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2">&nbsp;</td> 
                <%end if%> 
           <%objrst1.Close %>   
                
            </tr>
             
            <%objrst.Close %>
          
          <% objrst.Source = "select bztime,sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and kem='"& km &"' and bztime>='"& riqi &"' and bztime<='"& riqi1 &"' group by bztime"
            objRst.Open %>
           <%  do while not objrst.EOF%>
            <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year(objrst("bztime"))%>年<%=month(objrst("bztime"))%>月</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
           <% objrst1.Source = "select * from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea='"& objrst("bztime") &"'"
            objRst1.Open 
            if not objrst1.BOF  then
            if  objrst1("mon")<>0 then   %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
             
              <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                 <%else%>
               <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
           <%end if%>    
           <%else%>
               <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
           <%end if%>      
           <%objrst1.Close %>
          
            </tr>
           <%objrst.MoveNext
              loop %>
            <%objrst.Close  %> 
            <% objrst.Source = "select sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and kem='"& km &"' and bztime>='"& date1 &"' and bztime<='"& date2 &"' group by kem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年春季</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>='"& date1 &"' and yea<='"& date2 &"' group by kem"
            objRst1.Open 
            if  not objrst1.BOF   then
             if  objrst1("mon")<>0 then %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
                <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                   <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td>  
           <%end if%>
           <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td>  
           <%end if%>
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %> 
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  kem='"& km &"' and bztime>'"& date2 &"' and bztime<='"& date3 &"' group by kem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年夏季</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>'"& date2 &"' and yea<='"& date3 &"' group by kem"
            objRst1.Open 
            if not objrst1.BOF   then
              if  objrst1("mon")<>0 then %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>  
            <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %>
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  kem='"& km &"' and bztime>'"& date3 &"' and bztime<='"& date4 &"' group by kem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年秋季</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>'"& date3 &"' and yea<='"& date4 &"' group by kem"
            objRst1.Open 
             if not objrst1.BOF  then
             if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
             <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                    <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>   
             <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>      
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %>
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  kem='"& km &"' and bztime>'"& date4 &"' and bztime<='"& date5 &"' group by kem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年冬季</td>
                <td align="middle" ><font class="px12" color="black"><%=km%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& km &"' and yea>='"& date4 &"' and yea<='"& date5 &"' group by kem"
            objRst1.Open 
              if not objrst1.BOF  then
             if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
                 <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  '按科目方式进行上级核定数考核 结束  %>   
             </tbody></table></td></tr></tbody></table>
       <%else
       objrst.Close '按分科目方式进行上级核定数考核
        %>
       
         <%    
               objrst2.Source ="select distinct fkem from shenzhencwys_dep where dep='"& session("dep") &"' and kem='"& km &"'"
               objrst2.Open 
               do while not objrst2.EOF 
               objrst.Source = "select fkem,sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and fkem='"& objrst2("fkem") &"' and bztime>='"& riqi &"' and bztime<='"& riqi1 &"' group by fkem"
               objRst.Open 
               if not(objrst.BOF and objrst.EOF) then 
               
               %>
         
            <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年总计</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
             <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea>='"& riqi &"' and yea<='"& riqi1 &"' group by kem"
                objRst1.Open
                if not objrst1.BOF   then 
                 if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
              <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black" ><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                 <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2">&nbsp;</td> 
                <%end if%> 
                <%else%>
                 <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2">&nbsp;</td> 
                <%end if%> 
           <%objrst1.Close %>   
                
            </tr>
             
            <%objrst.Close %>
          
          <% objrst.Source = "select bztime,sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and fkem='"& objrst2("fkem") &"' and bztime>='"& riqi &"' and bztime<='"& riqi1 &"' group by bztime"
            objRst.Open %>
           <%  do while not objrst.EOF%>
            <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year(objrst("bztime"))%>年<%=month(objrst("bztime"))%>月</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
           <% objrst1.Source = "select * from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea='"& objrst("bztime") &"'"
            objRst1.Open 
            if not objrst1.BOF  then
            if  objrst1("mon")<>0 then   %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
              <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                 <%else%>
               <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
           <%end if%>    
           <%else%>
               <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
           <%end if%>      
           <%objrst1.Close %>
          
            </tr>
           <%objrst.MoveNext
              loop %>
            <%objrst.Close  %> 
            <% objrst.Source = "select sum(ysmoney)'m' from shenzhencwys  where dep='"& session("dep") &"' and fkem='"& objrst2("fkem") &"' and bztime>='"& date1 &"' and bztime<='"& date2 &"' group by fkem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年春季</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea>='"& date1 &"' and yea<='"& date2 &"' group by kem"
            objRst1.Open 
            if  not objrst1.BOF   then
             if  objrst1("mon")<>0 then %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td>  
           <%end if%>
           <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td>  
           <%end if%>
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %> 
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  fkem='"& objrst2("fkem") &"' and bztime>'"& date2 &"' and bztime<='"& date3 &"' group by fkem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年夏季</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea>'"& date2 &"' and yea<='"& date3 &"' group by kem"
            objRst1.Open 
            if not objrst1.BOF   then
              if  objrst1("mon")<>0 then %>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                 <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>  
            <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %>
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  fkem='"& objrst2("fkem") &"' and bztime>'"& date3 &"' and bztime<='"& date4 &"' group by fkem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年秋季</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea>'"& date3 &"' and yea<='"& date4 &"' group by kem"
            objRst1.Open 
             if not objrst1.BOF  then
             if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>   
             <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>      
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close  %>
              <% objrst.Source = "select sum(ysmoney)'m' from  shenzhencwys  where dep='"& session("dep") &"' and  fkem='"& objrst2("fkem") &"' and bztime>'"& date4 &"' and bztime<='"& date5 &"' group by fkem"
            objRst.Open 
            if  not objrst.BOF  then
            %>  
              <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year1%>年冬季</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst2("fkem")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("m")%></td>
               <% objrst1.Source = "select sum(mon)'mon' from shenzhencwys_je  where dep='"& session("dep") &"' and kem='"& objrst2("fkem") &"' and yea>='"& date4 &"' and yea<='"& date5 &"' group by kem"
            objRst1.Open 
              if not objrst1.BOF  then
             if  objrst1("mon")<>0 then%>      
                <td align="middle" ><font class="px12" color="black"><%=objrst1("mon")%></td>
               <%if objrst("m")/objrst1("mon")>=1 then%>
                <td align="middle" colspan="2" bgcolor=red><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                <%else%>
                 <td align="middle" colspan="2"><font class="px12" color="black"><%=formatpercent(objrst("m")/objrst1("mon"),1,0)%></td> 
                 <%end if%> 
                  <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
                 <%else%>
                <td align="middle" >&nbsp;</td>
                <td align="middle" colspan="2" >&nbsp;</td> 
            <%end if%>     
           <%objrst1.Close %>
           </tr> 
            <%end if%> 
             <%objrst.Close '按分科目方式进行上级核定数考核 结束 %>   
              <%
               else
               objrst.Close
               end if
               objrst2.MoveNext
              loop %>
             </tbody></table></td></tr></tbody></table>
             
       <%end if%>      
<%end if%>    
  
 </form>                 
   <br>
    
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
