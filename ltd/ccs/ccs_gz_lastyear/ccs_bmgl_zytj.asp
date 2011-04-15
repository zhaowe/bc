<%@ Language=VBScript %>

<% 
 
 'if trim(session("UID"))<>"" then
  ' dim objD
  ' set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
  '     VerifyOk=objD.VerifyUserFunction (session("UID"),"CCS_GSCXY")
  ' if VerifyOk=false then
  '    session("errorNo")="000002"
      'Response.Redirect "../sorry/sorry.asp"
 '  end if   
' else
  '  session("errorNo")="000001"
  '  Response.Redirect "../sorry/sorry.asp"
' end if 
 
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
     Response.Write "êy?Y±í3?′í￡?μ???è??ò2?μ?2???"
   else
     bm=objrst("companyname")
   end if
   
   objrst.Close %>
  
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
sql3="select * from companylocale where companyid='"&cid&"'"
rs_3.Open sql3,conn_1,3,3,1

f=trim(rs_3("companyname"))


'Response.Write(f)
'Response.Write(t)

set rs_k=server.CreateObject("adodb.recordset")                                                                            
rs_k.CursorLocation=2                                                                            
sql1="SELECT fkmshuom FROM cwys_km WHERE depar='"& f &"' and nian='"&year(date())&"' order by sn"                                                                            
rs_k.Open sql1,conn_1,3,3,1

km=Request.QueryString ("km")
mnykm=Request.QueryString("mnykm")
if km="" then
if mnykm="" then
km=trim(rs_k("fkmshuom"))
else
km=mnykm
end if
end if
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

  <%lb=Request.QueryString ("month")
  
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

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333ff">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/obj_hed1.gif">
      <table border="0" cellspacing="0" cellpadding="0" width="800">
        <tr>
          <td width="10"><img height="35" src="images/spacer.gif" width="10"></td>
          <td width="171"><img height="30" src="images/obj_maintitle.gif" width="171" border="0"></td>
          <td width="10"><img height="35" src="images/spacer.gif" width="10"></td>
          <td width="605"> 
            <table border="0" cellspacing="0" cellpadding="3" name="menubutton">
              <tr>
                <td><a href="ccs_bm_ser.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/ysglinfoin1.jpg',1)"><img src="images/ysglinfoin2.jpg" width="120" height="25" border="0" name="images1"></a></td>
              
                <td><a href="ccs_input_index.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/bt_02_off.gif',1)"><img src="images/bt_02_on.gif" width="120" height="25" border="0" name="images2"></a></td>
                <td><a href="ccs_bmgl_zttj.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/gltj2.jpg',1)"><img src="images/gltj1.jpg" width="120" height="24" border="0" name="images4"></a></td>
                <td><a href="ccs_bm_rzcx.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images3','','images/bmrz2.jpg',1)"><img src="images/bmrz1.jpg" width="120" height="24" border="0" name="images3"></a></td>
                <td><a href="ccs_bm_ytcx.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images5','','images/bmyt2.jpg',1)"><img src="images/bmyt1.jpg" width="120" height="24" border="0" name="images5"></a></td>
                <td><a href="ccs_bm_tscx.asp?km=<%=km%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images6','','images/tscx2.jpg',1)"><img src="images/tscx1.jpg" width="120" height="24" border="0" name="images6"></a></td>
              
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td bgcolor="#ff6600"><img height="2" src="images/spacer.gif" width="10"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="700">
  <tr> 
    <td width="140"><img height="98" src="images/obj_hed2_left.jpg" width="140"></td>
    <td width="30"><img height="98" src="images/obj_hed2_center2.jpg" width="30"></td>
    <td width="470" valign="top" background="images/obj_hed2_right.jpg"> 
      <table width="300" border="0" cellspacing="0" cellpadding="0" background>
        <tr>
          <td><img height="5" src="images/spacer.gif" width="330"></td>
          <td><img height="5" src="images/spacer.gif" width="130"></td>
        </tr>
        <tr>
          <td valign="top">
            <table border="0" cellspacing="0" cellpadding="0" name="banner">
              <tr> 
                <td colspan="2"><img height="24" src="images/spacer.gif" width="10"></td>
              </tr>
              <tr> 
                <td><img height="39" src="images/ba1_point.gif" width="35"></td>
                <td><img height="39" src="images/lizztp1.gif" width="205"></td>
              </tr>
            </table>
          </td>
          <td align="right">
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img height="53" src="images/spacer.gif" width="20"></td>
              </tr>
              <tr>
                <td><a href><img height="28" src="images/lizztp5.gif" width="138" border="0"></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
    </td>
    <td width="470" valign="top"><img height="73" src="images/obj_hed2_right-.jpg" width="60"> 
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
                <td><img height="45" src="images/spacer.gif" width="140"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="5" cellpadding="0" width="140">

       
         <tr>
  
            <tr> 
          <td> 
            <hr width="130">
          </td>
        </tr>
      
        <tr> 
          <td> 
            <table width="130" border="0" cellspacing="0" cellpadding="1">
                  <%objrst.Source ="select distinct sx=fkmcode,kem=fkmshuom,depar from cwys_km where depar='"& bm &"' and nian='"& syear &"' order by sx" 
                  'objrst.Source ="select distinct kem,sx from shenzhencwys_dep where dep='"& session("dep") &"' order by sx" 
                    objrst.Open 
                    if objrst.EOF and objrst.BOF then
                    else
                      objrst.MoveFirst
                      k=objrst("kem") 
                    
                                         
                      while not objrst.EOF %>
                     <tr> 
                     <td width="6" valign="top"><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>                     
                     <td width="116"><a href="ccs_bmgl_zytj.asp?syear=<%=syear%>&bm=<%=bm%>&km=<%=objrst("kem")%>&sfkmcode=<%=objrst("sx")%>"><font class="px14" color="#FFFFFF"><%=objrst("kem")%></font></a></td>
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
    
  </tr>   
      </table>
    </td>
    <td width="5"><img height="5" src="images/spacer.gif" width="10"></td>
    <td width="595" valign="top"> 
   
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="20" cellSpacing="0" cellPadding="0" width="775" border="0">
  <tbody>
  
  
 
      
   <tr>
    <td align="right">&nbsp;</td></tr>
   <tr>
    <td align="middle"><font class="px12" color="black">您目前进入的是<font color="red"><%=f%></font>&nbsp;<font color="blue"><%=km%></font>费用<font color="red"><font color="blue"><%="摘要查询"%></font><%=cxsj%></font>页面</font></td></tr>
    <tr>
    <td align="right">&nbsp;</td></tr>
     <tr>
    <td align="left">
    
    <table>
      <tr><td>
      
      
      <form action="ccs_bmgl_zytj_lb.asp?km=<%=km%>" method="post" id="form1" name="Form_Kmcode">
           
      <font class="px12" color="red"> 费用期间</font>
      <select name="sel_year" style="HEIGHT: 22px; WIDTH: 57px"> 
        <option selected value="<%=syear%>"><%=syear%></option>
        <option value="<%=year(date)-1%>"><%=year(date)-1%></option>
        <option value="<%=year(date)%>"><%=year(date)%></option>
        <option value="<%=year(date)+1%>"><%=year(date)+1%></option>
      </select>
      <font class="px12" color="red">年</font>&nbsp;
      <select name="sel_month_begin">         
        <option selected value="<%=1%>"><%=1%></option>
        <%for i=1 to 12 %>
        <option value="<%=i%>"><%=i%></option>        
        <%next %>
      </select>
      <font class="px12" color="red">月至</font>&nbsp;  
      <select name="sel_month_end">       
        <option selected value="<%=month(date)%>"><%=month(date)%></option>
        <%for i=1 to 12 %>
        <option value="<%=i%>"><%=i%></option>        
        <%next %>
      </select>      
      <font class="px12" color="red">月</font>&nbsp;
      
      <select name="selbm"> 
      
      <option value="<%=f%>"><%=f%></option>
  
        </select>      

      <font class="px12" color="red">科目代码</font>&nbsp;
 <%sfkmcode=Request.QueryString ("sfkmcode")%>  
  <input id="text1" name="Kmcode" style="WIDTH: 70px; HEIGHT: 21px" size="14" value="<%=sfkmcode%>">                     
      <font class="px12" color="red">费用摘要</font>&nbsp;                    
      <input id="text1" name="fyzy" style="WIDTH: 70px; HEIGHT: 21px" size="14">



      <input type="submit" value="确 定" id="submit1" name="submit1">

      </form>      
      
      
      
      </td>
          
          </tr>    
    </table>
    </td></tr>
   </tbody></table>
   
   

   
<table align="center" cellSpacing="0" cellPadding="0" width="777" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="73%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20">                    
         <font class="px12" color="white"> </font>
          </td>
        </tr>
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>

                

                
             
             </tbody></table></td></tr></tbody></table>
       
             
            

  
 
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
