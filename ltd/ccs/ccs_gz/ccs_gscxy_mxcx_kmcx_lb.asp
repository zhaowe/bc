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


  
  syear=Request.Form("sel_year") 
  month1=Request.Form("sel_month_begin")
  month2=Request.Form("sel_month_end")  
  kmcode1=Request.Form("kmcode_1")  
  kmcode2=Request.Form("kmcode_2")  
  
  
  
  
  
  %>
<html>
<head>
<title>预算管理系统--全公司查询</title>
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

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333ff" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/obj_hed1.gif">
      <table border="0" cellspacing="0" cellpadding="0" width="800">
        <tr>
          <td width="10"><IMG height=35 src="images/spacer.gif" width=10></td>
          <td width="171"><IMG height=30 src="images/obj_maintitle.gif" width=171 border=0></td>
          <td width="10"><IMG height=35 src="images/spacer.gif" width=10></td>
          <td width="605"> 
            <table border="0" cellspacing="0" cellpadding="3" name="menubutton">
              <tr>
                <td>&nbsp;</td>
                 <td>&nbsp;</td>
                <td>&nbsp;</td>
                 <td>&nbsp;</td>                 
                 <td><A onmouseover="MM_swapimages('images1','','images/bmcx.gif',1)" onmouseout=MM_swapImgRestore() href="ccs_gscxy_main.asp"><IMG height=20 src="images/bmcx_1.gif" width=100 border=0 name=images1></a></td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td><A onmouseover="MM_swapimages('images4','','images/fycx.gif',1)" onmouseout=MM_swapImgRestore() href="ccs_gscxy_index.asp"><IMG height=20 src="images/fycx_1.gif" width=100 border=0 name=images4></a></td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td>&nbsp;</td>
                 <td><a onmouseover="MM_swapimages('images4','','images/flcx_cf.jpg',1)" onmouseout="MM_swapImgRestore()" href="ccs_gscxy_mxcx.asp"><img height="20" src="images/flcx_cf.jpg" width="100" border="0" name="images4"></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td bgcolor="#ff6600"><IMG height=2 src="images/spacer.gif" width=10></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="700">
  <tr> 
    <td width="140"><IMG height=98 src="images/obj_hed2_left.jpg" width=140></td>
    <td width="30"><IMG height=98 src="images/obj_hed2_center2.jpg" width=30></td>
    <td width="470" valign="top" background="images/obj_hed2_right.jpg"> 
      <table width="300" border="0" cellspacing="0" cellpadding="0" background="">
        <tr>
          <td><IMG height=5 src="images/spacer.gif" width=330></td>
          <td><IMG height=5 src="images/spacer.gif" width=130></td>
        </tr>
        <tr>
          <td valign="top">
            <table border="0" cellspacing="0" cellpadding="0" name="banner">
              <tr> 
                <td colspan="2"><IMG height=24 src="images/spacer.gif" width=10></td>
              </tr>
              <tr> 
                <td><IMG height=39 src="images/ba1_point.gif" width=35></td>
                <td><IMG height=39 src="images/lizztp1.gif" width=205></td>
              </tr>
            </table>
          </td>
          <td align="right">
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><IMG height=53 src="images/spacer.gif" width=20></td>
              </tr>
              <tr>
                <td><a href=""><IMG height=28 src="images/lizztp5.gif" width=138 border=0></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
    </td>
    <td width="470" valign="top"><IMG height=73 src="images/obj_hed2_right-.jpg" width=60> 
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
                <td><IMG height=45 src="images/spacer.gif" width=140></td>
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
                <td width="6" valign="top"><img height="28" src="images/a.gif" width="28"></td>                     
                <td width="116"><a href="ccs_gscxy_mxcx_kmcx.asp"><font class="px14" color="#ffffff">科目查询</font></a></td>
             </tr>                            
             <tr> 
                <td width="6" valign="top"><img height="28" src="images/a.gif" width="28"></td>                     
                <td width="116"><a href="ccs_gscxy_mxcx_ytcx.asp"><font class="px14" color="#ffffff">预提查询</font></a></td>
             </tr> 
             <tr> 
                <td width="6" valign="top"><img height="28" src="images/a.gif" width="28"></td>                     
                <td width="116"><a href="ccs_gscxy_mxcx_tscx.asp"><font class="px14" color="#ffffff">托收查询</font></a></td>
             </tr>                                              
             <tr> 
                <td width="6" valign="top"><img height="28" src="images/a.gif" width="28"></td>                     
                <td width="116"><a href="ccs_gscxy_mxcx_zycx.asp"><font class="px14" color="#ffffff">摘要查询</font></a></td>
             </tr>            
             <tr> 
                <td width="6" valign="top"><img height="28" src="images/a.gif" width="28"></td>                     
                <td width="116"><a href="ccs_gscxy_mxcx_czcx.asp"><font class="px14" color="#ffffff">超支查询</font></a></td>
             </tr>                                                         
            
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
       
         <tr>
  
    
    
  </tr>   
      </table>
    </td>
    <td width="5"><IMG height=5 src="images/spacer.gif" width=10></td>
    <td width="595" valign="top"> 
   
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="20" cellSpacing="0" cellPadding="0" width="775"  border="0">
  <tbody>
  
  
 
      
   <tr>
    <td align="right" >&nbsp;</td></tr>
   <tr>
    <td align="middle" ><font class="px12" color="black">您目前进入的是<font color=red><font color=blue><%="费用科目查询"%></font><%=cxsj%></font></font></td></tr>
    
    
     <tr>
    <td align="left" >
    
    <table>
      <tr><td>
      
      
      <FORM action="ccs_gscxy_mxcx_kmcx_lb.asp" method=post id=form1 name="Form_Kmcode">
           
      <font class="px12" color="red"> 请选择费用期间</font>&nbsp;
      <select name="sel_year" style="HEIGHT: 22px; WIDTH: 57px"  > 
        <option selected value="<%=syear%>"><%=syear%></option>
        <option value="<%=year(date)-1%>"><%=year(date)-1%></option>
        <option value="<%=year(date)%>"><%=year(date)%></option>
        <option value="<%=year(date)+1%>"><%=year(date)+1%></option>
      </select>
      <font class="px12" color="red">年</font>&nbsp;
      <select name="sel_month_begin" >         
        <option selected value="<%=1%>"><%=1%></option>
        <%for i=1 to 12 %>
        <option value="<%=i%>"><%=i%></option>        
        <%next %>
      </select>
      <font class="px12" color="red">至</font>&nbsp;  
      <select name="sel_month_end" >       
        <option selected value="<%=month(date)%>"><%=month(date)%></option>
        <%for i=1 to 12 %>
        <option value="<%=i%>"><%=i%></option>        
        <%next %>
      </select>      
      <font class="px12" color="red">月</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        
      
      <font class="px12" color="red">请输入科目一级代码</font>&nbsp;
      <INPUT id=text1 name="Kmcode_1" style="WIDTH: 70px; HEIGHT: 21px" size=16>      
      <font class="px12" color="red">科目二级代码</font>&nbsp;                    
      <INPUT id=text1 name="Kmcode_2" style="WIDTH: 70px; HEIGHT: 21px" size=14>
            
      <INPUT type="submit" value="确 定" id=submit1 name=submit1>

      </FORM>        
      
      
      
      </td>
          
          </tr>    
    </table>
    </td></tr>
   </tbody></table>
   
   
   

    <%
   
  
sql_ed1=get_sql_ed(cstr(cint(month1)-1))
sql_ed2=get_sql_ed(cstr(month2))



if kmcode1="" then 
  sql_km1=" "
else
  sql_km1=" and kmcode like '" & kmcode1 &  "%' "
end if

if kmcode2="" then 
  sql_km2=" "
else
  sql_km2=" and fkmcode='" & kmcode2 &  "' "
end if



  
  
sql="  select ed.*,je=isnull(je,0),bl=isnull(je,0)/ze from "
sql=sql+" ( "
sql=sql+" select ysed.ys_year,ysed.ze,yskm.* from "
sql=sql+" ( "
sql=sql+" select ys_year,fkmcode,ze=sum(ze) from "
sql=sql+" ( "
sql=sql+" select ys_year,fkmcode," 

sql=sql+" ze=(("
'"jan+feb+mar+apr"
sql=sql+sql_ed2 
sql=sql+" )-("
sql=sql+sql_ed1  
'"jan+feb"
sql=sql+" )) "
'变量
sql=sql+" from cwys_ed where ys_year='" & syear & "' "
sql=sql+" ) as ysed "
sql=sql+" group by ys_year,fkmcode "
sql=sql+" ) as ysed "
sql=sql+" left join "
sql=sql+" ( "
sql=sql+" select distinct kmcode,kmshuom,fkmcode,fkmshuom from cwys_km where nian='" & syear & "' "

sql=sql+sql_km1+sql_km2
'sql=sql+" and kmcode='5408002' and fkmcode='811800' "
'变量
sql=sql+" ) as yskm "
sql=sql+" on ysed.fkmcode=yskm.fkmcode "
sql=sql+" where kmcode is not null and ysed.fkmcode is not null "

sql=sql+" ) as ed "
sql=sql+" left join " 
sql=sql+" ( "

sql=sql+" select mnykmcode,je=sum(price) from "
sql=sql+" ( "
sql=sql+" select mnykmcode,price from cwys_infoin  "
sql=sql+" where mnyyear='" & syear & "'  and month(mnytime)>=" & cint(month1) &" and month(mnytime)<=" & cint(month2) &" "
'变量
sql=sql+" and ifhandin='是' and cz<>'删除' "
sql=sql+" ) as fy "
sql=sql+" group by mnykmcode  "
sql=sql+" ) as fy "
sql=sql+" on ed.fkmcode=fy.mnykmcode "
sql=sql+" order by kmcode,fkmcode "

  
  
  
  sql_lb=sql

  session("kmcx_sql")=sql 
   
   
   
   'Response.Write sql_lb
   
   
   
   objrst.Source = sql_lb
   
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
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20" align="right">                    
          <A href="ccs_gscxy_mxcx_kmcx_lb_dc.asp"><font class="px12" color=white>导出EXCEL文件</font></A> 
          </td>
        </tr>
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#7dadc4" height="20">
                                                                                           
                <td align="middle"  ><font class="px12" color="black">科目一级代码</td>  
                <td align="middle"  ><font class="px12" color="black">科目名称</td>                                                              
                <td align="middle"  ><font class="px12" color="black">科目二级代码</td>                                  
                <td align="middle"  ><font class="px12" color="black">费用科目</td>                                
                <td align="middle"  ><font class="px12" color="black">指标</td>
                <td align="middle"  ><font class="px12" color="black">完成情况</td>
                <td align="middle"  ><font class="px12" color="black">剩余情况</td>
                <td align="middle"  ><font class="px12" color="black">完成率</td>                
                </tr>
                
            
                
                <%
                while not objrst.EOF 
                 cbys="#ecf7fd"
                 if objrst("bl")>1 and not isnumeric(lb)  then cbys="#ef867a"
                %> 
                <tr bgColor="<%=cbys%>" height="20">                
                <td align="middle" ><font class="px12" color="black"><%=objrst("kmcode")%></td>                
                <td align="middle" ><font class="px12" color="black"><%=objrst("kmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmcode")%></td>
                <td align="middle" ><font class="px12" color="black"><a href="ccs_gscxy_index.asp?syear=<%=syear%>&km=<%=trim(objrst("fkmshuom"))%>&month=<%=lb%>"><%=objrst("fkmshuom")%></A></td>                
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                <%
                objrst.MoveNext 
                wend%> 
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
  </tr></tbody>
</table></tr></table>
</body>
</html>

 <%
 function Get_Sql_ed(sel_month) 
   
   select case  cstr(sel_month)
     case "1":
           sql_ze="jan"  
     case "2":
           sql_ze="jan+feb"
     case "3":
           sql_ze="jan+feb+mar"
     case "4":
           sql_ze="jan+feb+mar+apr"
     case "5":
           sql_ze="jan+feb+mar+apr+may"
     case "6":
           sql_ze="jan+feb+mar+apr+may+jun"
     case "7":
           sql_ze="jan+feb+mar+apr+may+jun+jul"
     case "8":
           sql_ze="jan+feb+mar+apr+may+jun+jul+aug"
     case "9":
           sql_ze="jan+feb+mar+apr+may+jun+jul+aug+sep"
     case "10":
           sql_ze="jan+feb+mar+apr+may+jun+jul+aug+sep+oct"
     case "11":
           sql_ze="jan+feb+mar+apr+may+jun+jul+aug+sep+oct+nov"
     case "12":
           sql_ze="jan+feb+mar+apr+may+jun+jul+aug+sep+oct+nov+dece"                                           
     case else
           sql_ze="0"
   end select 
   Get_Sql_ed=sql_ze    
 
 end function 
%>