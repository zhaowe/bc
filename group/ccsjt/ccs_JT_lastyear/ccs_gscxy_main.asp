<%@ Language=VBScript %>

<% 
 
 if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"CCS_GSCXY")
   if VerifyOk=false then
      session("errorNo")="000002"
      'Response.Redirect "../sorry/sorry.asp"
   end if   
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
     bm=objrst("companyname")
   end if
   
   objrst.Close 
  
  'bm="信息工程部"
  
  'session("dep")=bm
  
  
  
  'year1=year(date)
  'date1=cdate(year1 & "-" & "1" & "-"&"1")
  'date2=cdate(year1 & "-" & "3" & "-"&"1")
  'date3=cdate(year1 & "-" & "6" & "-"&"1")
  'date4=cdate(year1 & "-" & "9" & "-"&"1")
  'date5=cdate(year1 & "-" & "12" & "-"&"1")
  'riqi=cdate(year1 & "-" & "1" & "-"&"1")
  'riqi1=cdate(year1 & "-" & "12" & "-"&"1")
  
  'km=Request.QueryString ("km")
  'session("km")=km

  lb=Request.QueryString ("month")
  
  syear=trim(Request.QueryString ("syear"))
  
  if syear="" then syear=year(date)
  
  %>
<html>
<head>
<title>深圳公司预算管理系统--全公司查询</title>
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

function year1_onchange() {
  syear=year1.value ;
  surl='ccs_gscxy_main.asp?syear='+syear+'&month=';
  window.document.location.href(surl);
}

function selkm_onchange() {
  sbm=selbm.value ;
  syear=year1.value ;
  //if (sbm=="所有科目")
  //   surl='ccs_gscxy_main.asp?syear='+syear
  //else   
     surl='ccs_gscxy_index.asp?syear='+syear+'&km='+sbm;
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
                  <%
                    'sql="select distinct sx=fkmcode,kem=fkmshuom,depar from cwys_km where depar='"& session("dep") &"' and nian='"& syear &"' order by sx" 
                    'sql="select distinct depart=companyname from companylocale"
                    sql="select distinct depart=companyname from cwysglbm where ysxs='1' order by companyname"
                    objrst.Source =sql
                  
                    objrst.Open 
                    if objrst.EOF and objrst.BOF then
                    else
                      objrst.MoveFirst
                      'k=objrst("kem") 
                    
                                         
                      while not objrst.EOF %>
                     <tr> 
                      <td width="6" valign="top"><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>                     
                     <td width="116"><a href="ccs_gscxy_bmldcx_main.asp?syear=<%=syear%>&bm=<%=objrst("depart")%>"><font class="px14" color="#FFFFFF"><%=objrst("depart")%></font></a></td>
                     </tr>
                     <%objrst.MoveNext %> 
                       <% wend 
                     end if  
                       %>
                       <%objrst.Close%>
                                
            
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
       
         <tr>
  
    <td align=center><a href="ccs_gscxy_main.asp?syear=<%=syear%>"><font class="px14" color="#FFFFFF">各部门统计</font></a></td>
    
  </tr>   
      </table>
    </td>
    <td width="5"><img src="images/spacer.gif" width="10" height="5"></td>
    <td width="595" valign="top"> 
   
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="20" cellSpacing="0" cellPadding="0" width="775"  border="0">
  <tbody>
  
  
  <%
  sql_ze="ze=sum(niandu)"
   sql_fy=""
   cxsj=cstr(syear)+"全年"
   
   select case  cstr(lb)
     case "1":
           sql_ze="ze=sum(jan)"
           sql_fy=" and month(mnytime)=1 "
           cxsj="一月份"
     case "2":
           sql_ze="ze=sum(feb)"
           sql_fy=" and month(mnytime)=2 "
           cxsj="二月份"
     case "3":
           sql_ze="ze=sum(mar)"
           sql_fy=" and month(mnytime)=3 "
           cxsj="三月份"
     case "4":
           sql_ze="ze=sum(apr)"
           sql_fy=" and month(mnytime)=4 "
           cxsj="四月份"
     case "5":
           sql_ze="ze=sum(may)"
           sql_fy=" and month(mnytime)=5 "
           cxsj="五月份"
     case "6":
           sql_ze="ze=sum(jun)"
           sql_fy=" and month(mnytime)=6 "
           cxsj="六月份"
     case "7":
           sql_ze="ze=sum(jul)"
           sql_fy=" and month(mnytime)=7 "
           cxsj="七月份"
     case "8":
           sql_ze="ze=sum(aug)"
           sql_fy=" and month(mnytime)=8 "
           cxsj="八月份"
     case "9":
           sql_ze="ze=sum(sep)"
           sql_fy=" and month(mnytime)=9 "
           cxsj="九月份"
     case "10":
           sql_ze="ze=sum(oct)"
           sql_fy=" and month(mnytime)=10 "
           cxsj="十月份"
     case "11":
           sql_ze="ze=sum(nov)"
           sql_fy=" and month(mnytime)=11 "
           cxsj="十一月份"
     case "12":
           sql_ze="ze=sum(dece)"
           sql_fy=" and month(mnytime)=12 "
           cxsj="十二月份"
           
     case "jd1":
           sql_ze="ze=sum(jan+feb+mar)"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj="截止一季度"     
     case "jd2":
           sql_ze="ze=sum(jan+feb+mar+apr+may+jun)"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj="截止二季度"     
     case "jd3":
           sql_ze="ze=sum(jan+feb+mar+apr+may+jun+jul+aug+sep)"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj="截止三季度"     
     case "jd4":
           sql_ze="ze=sum(jan+feb+mar+apr+may+jun+jul+aug+sep+oct+nov+dece)"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj="截止四季度" 
     case "全年":
           sql_ze="ze=sum(niandu)"
           sql_fy=""
           cxsj=cstr(syear)+"全年"                 
   end select 
  
  %>
      
   <tr>
    <td align="right" >&nbsp;</td></tr>
   <tr>
    <td align="center" ><font class="px12" color="black">您目前进入的是<font color=red><font color=blue><%="深圳公司"%></font><%=cxsj%></font>各部门预算完成情况页面</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>
     <tr>
    <td align="right" >
    
    <table>
      <tr><td>
      

      <font class="px12" color="red">请选择科目</font>&nbsp;          
     
      <select name="selbm"  LANGUAGE=javascript onchange="return selkm_onchange()"> 
      
      <option selected value="所有科目">所有科目</option>
      <option value="所有科目">所有科目</option>
      <% 
        sql="select distinct sx=fkmcode,kem=fkmshuom from cwys_km where nian='"& syear &"' order by kem" 
        objrst.Source =sql
        objrst.Open 
        while not objrst.EOF 
       
       %>
        
        <option value="<%=trim(objrst("kem"))%>"><%=trim(objrst("kem"))%></option>
      <%objrst.MoveNext 
        wend 
        objrst.Close 
         %>  
        </select>

      
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
   
   
   
   
   sql="select depar,ys_year,ze,je=isnull(wc,0),bl=isnull(wc,0)/ze from  "
   sql=sql+"( " 
   sql=sql+"select depar,ys_year, "   
   sql=sql+sql_ze
   sql=sql+"  from cwys_ed where  ys_year='" & syear &"' group by depar,ys_year "
   sql=sql+" ) as a left join "
   sql=sql+" ( "
   sql=sql+" select mnydepm,wc=sum(price),mnyyear "
   sql=sql+" from cwys_infoin where mnyyear='" & syear &"' and ifhandin='是' and cz<>'删除'  "
   sql=sql+sql_fy
   sql=sql+" group by mnydepm,mnyyear "
   sql=sql+" ) as b "
   sql=sql+" on a.depar=b.mnydepm "

   sql_sort=" order by ze desc "
   sql_lb="select * from ( "+sql+" ) as taball "+sql_sort   

   
   objrst.Source = sql_lb
   'Response.Write sql_lb
   objRst.Open 
   if objrst.EOF and objrst.BOF   then %>    
     <P> 
     <P align=center><font class="px12" color="black"><FONT color=crimson><STRONG><%=syear%>年无科目预算数据
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
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff"><b></b></font> <font color=white><a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=1"><font class="px12" color=white>1月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=2"><font class="px12" color=white>2月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=3"><font class="px12" color=white>3月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=4"><font class="px12" color=white>4月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=5"><font class="px12" color=white>5月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=6"><font class="px12" color=white>6月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=7"><font class="px12" color=white>7月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=8"><font class="px12" color=white>8月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=9"><font class="px12" color=white>9月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=10"><font class="px12" color=white>10月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=11"><font class="px12" color=white>11月</a> <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=12"><font class="px12" color=white>12月   </a>
          <a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=jd1"><font class="px12" color=white>   第一季度   </a><a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=jd2"><font class="px12" color=white>第二季度   </a><a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=jd3"><font class="px12" color=white>第三季度   </a><a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=jd4"><font class="px12" color=white>第四季度   </a><a href="ccs_gscxy_main.asp?syear=<%=syear%>&month=全年"><font class="px12" color=white>全年   </a></font></td></tr>
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#7dadc4" height="20">
                <td align="middle"  ><font class="px12" color="black">帐目时间</td>                                
                <td align="middle"  ><font class="px12" color="black">部门</td>                                
                <td align="middle"  ><font class="px12" color="black">指标</td>
                <td align="middle"  ><font class="px12" color="black">完成情况</td>
                <td align="middle"  ><font class="px12" color="black">剩余情况</td>
                <td align="middle"  ><font class="px12" color="black">完成率</td>                
                </tr>
                <%
                i=1
                while not objrst.EOF 
                 cbys="#ecf7fd"
                 if objrst("bl")>1 and not isnumeric(lb) then cbys="#ef867a"
                 %>
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle" ><font class="px12" color="black"><%=cxsj%></td>
                <td align="middle" ><font class="px12" color="black"><a href="ccs_gscxy_bmldcx_main.asp?syear=<%=syear%>&bm=<%=trim(objrst("depar"))%>&month=<%=lb%>"><%=objrst("depar")%></A></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                <%
                objrst.MoveNext 
                wend%> 
                <%objrst.Close %>
                
     <%
   sql_hj="select ze=sum(ze),je=isnull(sum(je),0),bl=isnull(sum(je)/sum(ze),0) from ( "  
   sql_hj=sql_hj+sql
   sql_hj=sql_hj+" ) as sy "
   
   objrst.Source =sql_hj
   objrst.Open    
     %>        
               
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="blue"><%=cxsj%>合计</td>
                <td align="middle" ><font class="px12" color="blue"><a href="ccs_gscxy_main.asp?syear=<%=syear%>">所有部门</A></td>
                <td align="middle" ><font class="px12" color="blue"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="blue"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="blue"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>     
           <%objrst.Close %>     
                
             
             </tbody></table></td></tr></tbody></table>
       
             
            
<%end if%>    
  
 
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
