<%@ Language=VBScript %>

<% 
 
if trim(session("UID"))<>"" then
  dim objD
  set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk_gsld=objD.VerifyUserFunction (session("UID"),"ccs_gscxy")
       VerifyOk_gsfz=objD.VerifyUserFunction (session("UID"),"cs_gsfzcx")
    'if VerifyOk_gsld=false then
      'session("errorNo")="000002"
      'Response.Redirect "../../sorry/sorry.asp"
    'else
    '   bm=Request.QueryString ("bm") 
    'end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
 
  'dep=Request.QueryString ("dep")




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

   



   emid=trim(session("emid"))
   loginid=trim(session("loginid"))

   sql="SELECT loginid,name,a.companyid,companyname FROM logininfo as a,companylocale as b "
   sql=sql+" where a.companyid=b.companyid and loginid='"& trim(session("loginid")) &"'"   
   objrst.open sql
   
   if objrst.eof and objrst.bof then
     Response.Write "数据表出错，登录人找不到部门"
   else
     depart=trim(objrst("companyname"))
   end if   
   objrst.Close 
  
   bm=trim(Request.QueryString ("bm"))      
    
    'if VerifyOk_gsld=true and bm<>"" then
       session("dep")=bm
   ' else
      ' bm=depart
     '  session("dep")=depart
   ' end if
   
   'if bm="公司领导" then 
    '  bm=Request.QueryString ("bm")
   'else
   '   session("dep")=bm
   'end if      
  
  'bm="信息工程部"
   
   
  
  bm=session("dep")
  
  'year1=year(date)
  'date1=cdate(year1 & "-" & "1" & "-"&"1")
  'date2=cdate(year1 & "-" & "3" & "-"&"1")
  'date3=cdate(year1 & "-" & "6" & "-"&"1")
  'date4=cdate(year1 & "-" & "9" & "-"&"1")
  'date5=cdate(year1 & "-" & "12" & "-"&"1")
  'riqi=cdate(year1 & "-" & "1" & "-"&"1")
  'riqi1=cdate(year1 & "-" & "12" & "-"&"1")
  
  km=Request.QueryString ("km")
  syear=trim(Request.QueryString ("syear"))
  lb=Request.QueryString ("smonth")
  
  session("km")=km

  if syear="" then syear=year(date)
  
  %>
<html>
<head>
<title>深圳公司预算管理系统--<%=bm%></title>
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
a:active {  color: #0000;; text-decoration: none}
a:visited {  color: #000000; font-weight: normal;; text-decoration: none}
a:link {  color: #000000; font-weight: normal; ; text-decoration: none}
a.homepage:link {  color: #000000; font-weight: normal;}
a.homepage:visited {  color: #000000; font-weight: normal;}
a.homepage:active {  color:#000000 ; font-weight: normal;}
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

function year1_onchange() {
  var sbm;
  sbm='<%=bm%>' ;
  syear=year1.value ;
  surl='ccs_gscxy_bmldcx_index.asp?syear='+syear+'&bm='+sbm;
  window.document.location.href(surl);
}
function selbm_onchange() {
  sbm=selbm.value ;
  syear=year1.value ;
  if (sbm=="所有部门")
     surl='ccs_gscxy_main.asp?syear='+syear+'&bm='+sbm;
  else   
     surl='ccs_bmldcx_main.asp?syear='+syear+'&bm='+sbm;
  window.document.location.href(surl);
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333FF" >
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
                <td>&nbsp;</td>
                 <td>&nbsp;</td>                 
                 <td><a href="ccs_gscxy_main.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/bmcx.gif',1)"><img src="images/bmcx_1.gif" width="100" height="20" border="0" name="images1"></a></td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td><a href="ccs_gscxy_index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/fycx.gif',1)"><img src="images/fycx_1.gif" width="100" height="20" border="0" name="images4"></a></td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td><a href="ccs_gscxy_mxcx.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/flcx_cf.jpg',1)"><img src="images/flcx_cf.jpg" width="100" height="20" border="0" name="images4"></a></td>
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
                  <%objrst.Source ="select distinct sx=fkmcode,kem=fkmshuom,depar from cwys_km where depar='"& bm &"' and nian='"& syear &"' order by sx" 
                    objrst.Open 
                    if objrst.EOF and objrst.BOF then
                      
                    else
                      objrst.MoveFirst
                      k=objrst("kem") 
                      while not objrst.EOF %>
                     <tr> 
                      <td width="6" valign="top"><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                     <td width="116"><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=objrst("kem")%>"><font class="px14" color="#FFFFFF"><%=objrst("kem")%></font></a></td>
                     </tr>
                     <%objrst.MoveNext %> 
                       <% 
                      wend 
                    end if  
                       objrst.Close%>
                                
            
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
  <tr>
  
    <td align=center><a href="ccs_bmldcx_main.asp?syear=<%=syear%>&bm=<%=bm%>"><font class="px14" color="#FFFFFF">各项统计</font></a></td>
    
  </tr>         
       
       
      </table>
    </td>
    <td width="5"><img src="images/spacer.gif" width="10" height="5"></td>
    <td width="595" valign="top"> 
   
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="20" cellSpacing="0" cellPadding="0" width="775"  border="0">
  <tbody>
      
   <tr>
    <td align="right" >&nbsp;</td></tr>
   <tr>
   
   <%
   
   djrq=trim(cstr(syear))+"年"
   
   'Response.Write cstr(trim(lb))
   
   select case  cstr(trim(lb))
     case "": djyf=""
     case "jd1": djyf="截止一季度"
     case "jd2": djyf="截止二季度"
     case "jd3": djyf="截止三季度"
     case "jd4": djyf="截止四季度"
     case "全年": djyf=""
     case else:   djyf=cstr(trim(lb))+"月"
   end select      
   
   'Response.Write djyf
   
   djrq=djrq+djyf
   
   'Response.Write djrq
   
   %>
   
    <td align="center" ><font class="px12" color="black">您目前进入的是<font color=blue><%=bm%></font><%=djrq%><font color=red><%=km%></font>科目预算完成情况页面</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>
     <tr>
    <td align="right" >
    
    <table>
    
      <tr><td>
      
<% if VerifyOk_gsld=true then %>
      <font class="px12" color="red">请选择部门</font>&nbsp;          
     
      <select name="selbm"  LANGUAGE=javascript onchange="return selbm_onchange()"> 
      <option selected value="<%=bm%>"><%=bm%></option>
      <option value="所有部门">所有部门</option>
      <% 
        'sql="select distinct depart=companyname from companylocale order by companyname"
        sql="select distinct depart=companyname from cwysglbm where ysxs='1' order by companyname"
        objrst.Source =sql
        objrst.Open 
        while not objrst.EOF 
       
       %>
        
        <option value="<%=trim(objrst("depart"))%>"><%=trim(objrst("depart"))%></option>
      <%objrst.MoveNext 
        wend 
        objrst.Close 
         %>  
        </select>
<% end if  %>        

<% if VerifyOk_gsfz=true then %>
      <font class="px12" color="red">请选择部门</font>&nbsp;          
     
      <select name="selbm"  LANGUAGE=javascript onchange="return selbm_onchange()"> 
      <option selected value="<%=bm%>"><%=bm%></option>
      <option value="所有部门">所有部门</option>
      <% 
        'sql="select distinct depart=companyname from companylocale order by companyname"
        'sql="select distinct depart=companyname from cwysglbm where ysxs='1' order by companyname"
        sql="select distinct depart=companyname from cwysglbm where ysxs='1' and ysfgld='"& emid &"' order by companyname"
        objrst.Source =sql
        objrst.Open 
        while not objrst.EOF 
       
       %>
        
        <option value="<%=trim(objrst("depart"))%>"><%=trim(objrst("depart"))%></option>
      <%objrst.MoveNext 
        wend 
        objrst.Close 
         %>  
        </select>
<% end if  %> 
      
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      
      <font class="px12" color="black">  <font color="red">请选择财务预算年份&nbsp;      
      <select name="year1" style="HEIGHT: 22px; WIDTH: 57px"  LANGUAGE=javascript onchange="return year1_onchange()"> 
        <option selected value="<%=syear%>"><%=syear%></option>
        <option value="<%=year(date)-1%>"><%=year(date)-1%></option>
        <option value="<%=year(date)%>"><%=year(date)%></option>
        <option value="<%=year(date)+1%>"><%=year(date)+1%></option>
        </select>
      <font color="#ff0000"></font>
      </td>
          
          </tr>    
    </table>
    </td></tr>
   </tbody></table>
   
   
   <% 
   
   'if trim(lb)="" then 
   '  sql_fy=""
   'else
   '  sql_fy=" and month(mnytime)='"& lb &"' "
   'end if  
   
   select case  cstr(trim(lb))
     case "": sql_fy=""
     case "jd1": sql_fy=" and month(mnytime)>=1 and month(mnytime)<=3"
     case "jd2": sql_fy=" and month(mnytime)>=4 and month(mnytime)<=6"
     case "jd3": sql_fy=" and month(mnytime)>=7 and month(mnytime)<=9"
     case "jd4": sql_fy=" and month(mnytime)>=10 and month(mnytime)<=12"
     case "全年": sql_fy=""
     case else:  sql_fy=" and month(mnytime)='"& lb &"' "
   end select     
   
   
   
   sql="select * from cwys_infoin where mnydepm='"& bm &"' and mnykm='" & km &"' and ifhandin='是'  and cz<>'删除' "
   sql=sql+sql_fy+ " and mnyyear='"& syear &"' order by mnytime desc" 
   objrst.Source = sql
   'Response.Write sql
   objRst.Open 
%>
   
<table align="center"  cellSpacing="0" cellPadding="0" width="777" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="73%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff"><b></b></font> <font color=white><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=1"><font class="px12" color=white>1月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=2"><font class="px12" color=white>2月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=3"><font class="px12" color=white>3月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=4"><font class="px12" color=white>4月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=5"><font class="px12" color=white>5月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=6"><font class="px12" color=white>6月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=7"><font class="px12" color=white>7月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=8"><font class="px12" color=white>8月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=9"><font class="px12" color=white>9月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=10"><font class="px12" color=white>10月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=11"><font class="px12" color=white>11月</a> <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=12"><font class="px12" color=white>12月</a>
          <a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=jd1"><font class="px12" color=white>   第一季度   </a><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=jd2"><font class="px12" color=white>第二季度   </a><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=jd3"><font class="px12" color=white>第三季度   </a><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=jd4"><font class="px12" color=white>第四季度   </a><a href="ccs_gscxy_bmldcx_index.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=km%>&smonth=全年"><font class="px12" color=white>全年   </a></font></td></tr>
        <tr>
        
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
         <% if objrst.EOF and objrst.BOF   then %>    
          <tr> 
           <td>
            <P align=center><font class="px12" color="black"><FONT color=crimson><STRONG><%=km%>科目没有报帐信息
           </td> 
          </tr>
         <%else%>
       
               <tr bgColor="#7dadc4" height="20">
                <td align="middle"  ><font class="px12" color="black">帐目时间</td>                                
                <td align="middle"  ><font class="px12" color="black">费用科目</td>                                
                <td align="middle"  ><font class="px12" color="black">经办人</td>
                <td align="middle"  ><font class="px12" color="black">报销人</td>
                <td align="middle"  ><font class="px12" color="black">登记日</td>
                <td align="middle"  ><font class="px12" color="black">核销日</td>                
                <td align="middle"  ><font class="px12" color="black">金额</td>
                <td align="middle"  ><font class="px12" color="black">费用说明</td>
                <td align="middle"  ><font class="px12" color="black">付款方式</td>                
                <td align="middle"  ><font class="px12" color="black">录入人</td>
                </tr>
                <%
                i=1
                while not objrst.EOF 	
                 cbys="#ecf7fd"                 
                 if trim(objrst("isover"))="是" then cbys="#ef867a"
                 'if trim(objrst("isover"))="是" then cbys="red"                                
                %>
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=year(objrst("mnytime"))%>年<%=month(objrst("mnytime"))%>月</td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("mnykm")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("passname")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("bxname")%></td>                                                                
                <td align="middle" ><font class="px12" color="black"><%=objrst("djdate")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("hxdate")%></td>                
                <td align="middle"  ><font class="px12" color="black"><%=objrst("price")%></td>                
                <td align="middle"  ><font class="px12" color="black"><%=objrst("mnynote")%></td>
                <td align="middle"  ><font class="px12" color="black"><%=objrst("payway")%></td>
                <td align="middle"  ><font class="px12" color="black"><%=objrst("djname")%></td>                
                </tr>
                <%
                'i=i+1
                objrst.MoveNext 
                wend%> 
                
                <%end if%>            
                
                <%objrst.Close %>
         
             </tbody></table></td></tr>
             
             
<%
   
   sql_ze_nd="ze=a.niandu,"
   sql_fy_nd=""
   cxsj_nd=cstr(syear)+"全年"
   
   
   sql_ze_dyf=""
   sql_fy_dyf=""
   cxsj_dyf=""     
   
   
   if lb="" then lb=month(date)
   
   select case  cstr(lb)
     case "1":
           sql_ze_yf="ze=a.jan, "
           sql_fy_yf=" and month(mnytime)=1 "
           cxsj_yf="一月份"
           
           sql_ze_jd="ze=a.jan+a.feb+a.mar,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "           
           cxsj_jd="截止一季度"           
           
           sql_ze_dyf="ze=a.jan, "
           sql_fy_dyf=" and month(mnytime)=1 "
           cxsj_dyf="截止一月份"
           
     case "2":
           sql_ze_yf="ze=a.feb,"
           sql_fy_yf=" and month(mnytime)=2 "
           cxsj_yf="二月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_jd="截止一季度"
           
           sql_ze_dyf="ze=a.jan+a.feb, "
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=2 ) "
           cxsj_dyf="截止二月份"                     
                    
     case "3":
           sql_ze_yf="ze=a.mar,"
           sql_fy_yf=" and month(mnytime)=3 "
           cxsj_yf="三月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_jd="截止一季度"    
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_dyf="截止三月份"            
                    
     case "4":
           sql_ze_yf="ze=a.apr,"
           sql_fy_yf=" and month(mnytime)=4 "
           cxsj_yf="四月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_jd="截止二季度"     
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=4 ) "
           cxsj_dyf="截止四月份"                  
     case "5":
           sql_ze_yf="ze=a.may,"
           sql_fy_yf=" and month(mnytime)=5 "
           cxsj_yf="五月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_jd="截止二季度"     

           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=5 ) "
           cxsj_dyf="截止五月份"                    
     case "6":
           sql_ze_yf="ze=a.jun,"
           sql_fy_yf=" and month(mnytime)=6 "
           cxsj_yf="六月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_jd="截止二季度"         
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_dyf="截止六月份"            
               
     case "7":
           sql_ze_yf="ze=a.jul,"
           sql_fy_yf=" and month(mnytime)=7 "
           cxsj_yf="七月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_jd="截止三季度"       
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=7 ) "
           cxsj_dyf="截止七月份"                 
     case "8":
           sql_ze_yf="ze=a.aug,"
           sql_fy_yf=" and month(mnytime)=8 "
           cxsj_yf="八月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_jd="截止三季度"   
                      
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=8 ) "
           cxsj_dyf="截止八月份"                     
     case "9":
           sql_ze_yf="ze=a.sep,"
           sql_fy_yf=" and month(mnytime)=9 "
           cxsj_yf="九月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_jd="截止三季度"   
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_dyf="截止九月份"             
     case "10":
           sql_ze_yf="ze=a.oct,"
           sql_fy_yf=" and month(mnytime)=10 "
           cxsj_yf="十月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_jd="截止四季度"   
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.aug+a.sep+a.oct,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=10 ) "
           cxsj_dyf="截止十月份"                     
     case "11":
           sql_ze_yf="ze=a.nov,"
           sql_fy_yf=" and month(mnytime)=11 "
           cxsj_yf="十一月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_jd="截止四季度"   
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=11 ) "
           cxsj_dyf="截止十一月份"                     
     case "12":
           sql_ze_yf="ze=a.dece,"
           sql_fy_yf=" and month(mnytime)=12 "
           cxsj_yf="十二月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_jd="第四季度"
           
           sql_ze_dyf="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_dyf=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_dyf="截止十二月份"  
                      
           
     case "jd1":
           sql_ze_yf="ze=a.jan+a.feb+a.mar,"
           sql_fy_yf=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_yf="第一季度"
           sql_ze_jd="ze=a.jan+a.feb+a.mar,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_jd="截止一季度" 
     case "jd2":
           sql_ze_yf="ze=a.apr+a.may+a.jun,"
           sql_fy_yf=" and (month(mnytime)>=4 and month(mnytime)<=6 ) "
           cxsj_yf="第二季度"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_jd="截止二季度" 
     case "jd3":
           sql_ze_yf="ze=a.jul+a.aug+a.sep,"
           sql_fy_yf=" and (month(mnytime)>=7 and month(mnytime)<=9 ) "
           cxsj_yf="第三季度"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_jd="截止三季度"
     case "jd4":
           sql_ze_yf="ze=a.oct+a.nov+a.dece,"
           sql_fy_yf=" and (month(mnytime)>=10 and month(mnytime)<=12 ) "
           cxsj_yf="十二月份"
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_jd="截止四季度"
     case "全年":
           sql_ze_yf="ze=a.niandu,"
           sql_fy_yf=""
           cxsj_yf=cstr(syear)+"全年"      
           sql_ze_jd="ze=a.niandu,"
           sql_fy_jd=""
           cxsj_jd=cstr(syear)+"全年" 
     
   end select 
   
   
   
   
   
   
   sql="select ys_year,fkmshuom,ze,je=sum(je),bl=sum(je)/ze from " 
   sql=sql+"( " 
   sql=sql+"select a.*,je=isnull(b.price,0) from  " 
   sql=sql+"( " 
   sql=sql+"select a.depar,a.ys_year,a.fkmcode,"
   
   sql_p1=sql
   
   
   'sql=sql+"a.niandu,"
   'sql=sql+sql_ze
   
   sql=" b.fkmshuom from cwys_ed as a,cwys_km as b  " 
   sql=sql+" where a.depar='"& bm &"' and a.ys_year='"& syear &"' and a.depar=b.depar and a.fkmcode=b.fkmcode and a.ys_year=b.nian "
   
   sql_p2=sql
   
   sql_km=" and b.fkmshuom='"& km &"' " 
   
   sql=" ) as a " 
   sql=sql+"left join cwys_infoin b on  a.fkmcode=b.mnykmcode and a.depar=b.mnydepm and a.ys_year=b.mnyyear and ifhandin='是' and cz<>'删除'  "
   
   sql_p3=sql
   
   'sql=sql+sql_fy 
   
   sql=" ) as c " 
   sql=sql+"group by ys_year,fkmshuom,fkmshuom,ze " 
   
   sql_p4=sql  
   
%> 

                <%
                
                sql_hj_fy=sql_p1+sql_ze_yf+sql_p2+sql_km+sql_p3+sql_fy_yf+sql_p4

                 'Response.Write sql_hj_fy
   
                  objrst.Source = sql_hj_fy
   
                 objRst.Open                                  
                %>
                
         <% if objrst.EOF and objrst.BOF   then %>    
          <tr> 
           <td>
            <P align=center><font class="px12" color="black"><FONT color=crimson><STRONG><%=km%>科目<%=syear%>年未分配额度
           </td> 
          </tr>
         <%else%>             
             
             <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#7dadc4" height="20">
                <td align="middle"  ><font class="px12" color="black">帐目时间</td>                                
                <td align="middle"  ><font class="px12" color="black">费用科目</td>                                
                <td align="middle"  ><font class="px12" color="black">指标</td>
                <td align="middle"  ><font class="px12" color="black">完成情况</td>
                <td align="middle"  ><font class="px12" color="black">剩余情况</td>
                <td align="middle"  ><font class="px12" color="black">完成率</td>                
                </tr>
               
                
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj_yf%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>
                
                
                
                <%
                '显示截止月份
                if cxsj_dyf<>"" then                
            
                sql_hj_jd=sql_p1+sql_ze_dyf+sql_p2+sql_km+sql_p3+sql_fy_dyf+sql_p4
                 'Response.Write sql_hj_jd   
                  objrst.Source = sql_hj_jd   
                 objRst.Open                                  
                 cbys="#ecf7fd"
                 'if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">                                 
                <td align="middle" ><font class="px12" color="black"><%=cxsj_dyf%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>   
                </tr>                
                <%objrst.Close 
                
                end if
                %>                    
                
                
                
                
                <%
                
                

           sql_ze_jd="ze=a.jan+a.feb+a.mar,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj_jd="截止一季度"                 
                
                
                sql_hj_jd=sql_p1+sql_ze_jd+sql_p2+sql_km+sql_p3+sql_fy_jd+sql_p4

                 'Response.Write sql_hj_jd
   
                  objrst.Source = sql_hj_jd
   
                 objRst.Open                 
                 
                 cbys="#ecf7fd"
                 if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">                                 
                <td align="middle" ><font class="px12" color="black"><%=cxsj_jd%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>        
                <%
                
                
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj_jd="截止二季度"                
                
                
                sql_hj_jd=sql_p1+sql_ze_jd+sql_p2+sql_km+sql_p3+sql_fy_jd+sql_p4

                 'Response.Write sql_hj_jd
   
                  objrst.Source = sql_hj_jd
   
                 objRst.Open                 
                 cbys="#ecf7fd"
                 if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj_jd%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>   
                <%
                
                
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj_jd="截止三季度"               
                
                
                sql_hj_jd=sql_p1+sql_ze_jd+sql_p2+sql_km+sql_p3+sql_fy_jd+sql_p4

                 'Response.Write sql_hj_jd
   
                  objrst.Source = sql_hj_jd
   
                 objRst.Open                 
                 cbys="#ecf7fd"
                 if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj_jd%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>   
                <%
                
                
           sql_ze_jd="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy_jd=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj_jd="截止四季度"              
                
                
                sql_hj_jd=sql_p1+sql_ze_jd+sql_p2+sql_km+sql_p3+sql_fy_jd+sql_p4

                 'Response.Write sql_hj_jd
   
                  objrst.Source = sql_hj_jd
   
                 objRst.Open                 
                 cbys="#ecf7fd"
                 if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj_jd%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>                                                     
                
                <%
                
                sql_hj_nd=sql_p1+sql_ze_nd+sql_p2+sql_km+sql_p3+sql_fy_nd+sql_p4

                ' Response.Write sql_hj_nd
   
                  objrst.Source = sql_hj_nd
   
                 objRst.Open                 
                 cbys="#ecf7fd"
                 if objrst("bl")>1   then cbys="#ef867a"                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj_nd%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                
                <%objrst.Close %>                         
                
             
             </tbody></table></td></tr>
             
             
             <%end if%>                 
             </tbody></table>
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
