

<%@ Language=VBScript %>

<!--#include file="author1.asp" -->
  
<html>
<head>
<title>公司预算管理系统</title>
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
sql1="SELECT fkmshuom,fkmcode,kmshuom FROM cwys_km WHERE depar='"& f &"' and nian='"&year(date())&"' order by sn"                                                                            
rs_k.Open sql1,conn_1,3,3,1
if rs_k.EOF then
Response.Write("本年度您所在部门的各项额度尚未录入，请与财务部门联系！")
else


km=Request.QueryString ("km")
if km="" then
km=trim(rs_k("fkmshuom"))
end if



fkmcode1=Request.QueryString ("km1")
if fkmcode1="" then
fkmcode1=trim(rs_k("fkmcode"))
end if

km2=Request.QueryString ("km2")
if km2="" then
km2=trim(rs_k("kmshuom"))
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



function chk_form(d)
{


if (d.hbh.value=="")
{alert("经办人员工号不能为空")
d.hbh.focus();
return false}

if (d.hd.value=="")
{alert("经办人姓名不能为空")
d.hd.focus();
return false}

if (d.price.value=="")
{alert("金额不能为空")
d.price.focus();
return false}
 


if (confirm("你确定所选科目正确吗？")==false)

  return false

}


function ShowFLT(i) 
{
    lbmc = eval('LM' + i);
    
    if (lbmc.style.display == 'none') 
    {
        //LMYC();
        
        lbmc.style.display = '';
    }
    else 
    {
        
        lbmc.style.display = 'none';
    }
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

                <td><a href="ccs_bm_ser.asp?km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/ysglinfoin1.jpg',1)"><img src="images/ysglinfoin2.jpg" width="120" height="25" border="0" name="images1"></a></td>
              
                <td><a href="ccs_input_index.asp?km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/bt_02_off.gif',1)"><img src="images/bt_02_on.gif" width="120" height="25" border="0" name="images2"></a></td>
                <td><a href="ccs_bmgl_zttj.asp?km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/gltj2.jpg',1)"><img src="images/gltj1.jpg" width="120" height="24" border="0" name="images4"></a></td>
                <td><a href="ccs_bm_rzcx.asp?km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images3','','images/bmrz2.jpg',1)"><img src="images/bmrz1.jpg" width="120" height="24" border="0" name="images3"></a></td>
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
            
 
  
                     
            
            <TBODY>
                           <%
                           
                     Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst1=server.CreateObject ("ADODB.Recordset")
  objRst1.LockType=3
  objRst1.CursorType=3
  set objRst1.activeConnection=objConn
                  ' objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '信息工程部' and kmshuom='其他业务支出-通讯费' order by kmcode,kmshuom" 
                  objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '"&f&"' and nian='"&year(date())&"' order by kmcode,kmshuom" 
                 ' objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '"&f&"' order by kmcode,kmshuom" 

    
                       'Response.Write(objrst1.Source)
                     
                       objrst1.Open 
                       j=1
                       while not objrst1.EOF
                  
                   %>
                       
          <TR>
       
          
            <TD style="PADDING-LEFT: 0px" background="">
            <A onclick=javascript:ShowFLT(<%=j%>) 
                  href="javascript:void(null)"><font class="px14" color="#FFFFFF">+<%=trim(objrst1("kmshuom"))%></font></A> 
                  </TD>
                 
          </TR>
   
          <TR id=LM<%=j%> style="DISPLAY: none">
            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                 
               
                   <%
                     Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
                   'objrst.Source ="select distinct kmshuom,kmcode,fkmshuom,fkmcode from cwys_km where depar = '"&f&"' and kmshuom='"&trim(objrst1("kmshuom"))&"' order by kmcode,kmshuom,fkmcode,fkmshuom" 
                       'Response.Write(objrst.Source)
                   objrst.Source ="select distinct kmshuom,kmcode,fkmshuom,fkmcode from cwys_km where depar = '"&f&"' and kmshuom='"&trim(objrst1("kmshuom"))&"' and nian='"&year(date())&"' order by kmcode,kmshuom,fkmcode,fkmshuom" 
                       objrst.Open 
                       while not objrst.EOF
                   
                   %>

                  <TR>
                    <TD style="PADDING-LEFT: 0px" height=23> <A title=资料册 
                        href="ccs_input_index.asp?km=<%=trim(objrst("fkmshuom"))%>&km1=<%=trim(objrst("fkmcode"))%>&km2=<%=trim(objrst("kmshuom"))%>" 
                        ><font class="px12" color=white>*<%=trim(objrst("fkmshuom"))%></font></A> </TD>
                  </TR>
                  <TR>
                    <TD background="" height=3></TD>
                  </TR>
                                
           
                  
                    <%objrst.MoveNext %> 
                       <% wend %>    
                 
                    <%objrst.Close%>
                    
                  
                </TBODY>
            </TABLE></TD>
          </TR>
          
            <%
            objrst1.MoveNext 
            j=j+1
              %> 
                       <% wend %>    
                 
                    <%objrst1.Close%>
          
          
          
          
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
    <form method="post" action="ccs_input_index.asp?todo=02&km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" id=form1 name=form1>   
     <table style="BORDER-RIGHT: #4e4c71 1px solid; BORDER-TOP: #4e4c71 1px solid; BORDER-LEFT: #4e4c71 1px solid; BORDER-BOTTOM: #4e4c71 1px solid" height="92%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4e4c71" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff" class="px14">&nbsp;帐单录入&nbsp;&nbsp;/&nbsp;&nbsp;<a href="http://www.boc.cn/cn/common/whpj.html" target=_blank><font color="white">外汇牌价</a></font> </td></tr>
        <tr>
          <td vAlign="top" width="100%">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
              
            
   <tr>
    <td align="center" ><font class="px12" color="black">您目前进入的是&nbsp;<font color=red><%=km2%>(<%=km%>)</font><font color=blue> 账单录入</font>界面</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>

<tr width=100%>
     <td  align="left">
     <font class="px12" color="black">&nbsp;&nbsp;经办人员工号：</font>
     <input id="hbh" type="text" name="hbh" size="5" onchange="javascript:disname()"><font class="px12" color="black">*  </font>  
    
     <font class="px12" color="black">&nbsp;&nbsp;经办人姓名：</font>
     <input id="hd" type="text" name="hd" size="10"><font class="px12" color="black">*</font>
     
     <font class="px12" color="black">&nbsp;&nbsp;费用控制部门：</font>
     <input type="text" name="mnydepm" size="10" readonly value=<%=f%>>
<br>
 
     <font class="px12" color="black">&nbsp;&nbsp;报销人员工号：</font>
     <input type="text" name="hbh1" size="5" onchange="javascript:disname1()"> 
     
     <font class="px12" color="black">&nbsp;&nbsp;报销人姓名：</font>
     <input type="text" name="hd1" size="10" >
     
     <font class="px12" color="black">&nbsp;&nbsp;费用科目：</font>
     <input type="text" name="mnykm" size="20" readonly value=<%=km%>>

<font class="px12"><%=fkmcode1%></font>      
<br>
   
      </td>
    </tr>


    <tr>
      <td align="left">
    <font class="px12" color="black">&nbsp;&nbsp;费用期间：</font>
        <select size="1" name="year1">
         <OPTION  value="<%=cstr(year(date)+1)%>"><font color=black class="px14"><%=cstr(year(date)+1)%></OPTION>
             <OPTION selected value="<%=cstr(year(date))%>"><font color=black class="px14"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><font color=black class="px14"><%=cstr(year(date)-1)%></OPTION>
                       <OPTION  value="<%=cstr(year(date)-2)%>"><font color=black class="px14"><%=cstr(year(date)-2)%></OPTION>

            </select>
            <font color=black class="px14">年</font>
        <select name="month1" style="HEIGHT: 22px; WIDTH: 43px"> <option selected><%=month(date)%></option>
        <%dd=month(date())
        
          cha=dd-1
          while cha>0 
    %>
        <option value="<%=cha%>"><%=cha%></option>
        
        <%
        cha=cha-1
        wend%></select>
        
      <font class="px12" color="black">月</font>
     
     <font class="px12" color="black">&nbsp;&nbsp;付款方式：</font>
        <select size="1" name="payway">
         <OPTION  value="现金"><font color=black class="px14">现金</OPTION>
         <OPTION  value="银行"><font color=black class="px14">银行</OPTION>
         <OPTION  value="内部往来"><font color=black class="px14">内部往来</OPTION>
          <OPTION  value="转帐"><font color=black class="px14">转帐</OPTION>
            </select>
                
     <font class="px12" color="black"><br>&nbsp;&nbsp;金额：</font>
     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="price" size="10" ><font class="px12" color="black">元*</font>
     <font class="px12" color="black">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;帐单号：</font>
     <input type="text" name="tabid" size="15" ><font class="px12" color="black"></font>
    </td>
    
    </tr>
     <tr>
     <td  align="left">
     <font class="px12" color="black">&nbsp;&nbsp;费用说明：</font>
               <input type="text" name="mnynote" size="80" > 
     <br>
         <font class="px12" color="black">&nbsp;&nbsp;登记人：<%=name%>&nbsp;&nbsp;登记日：<%=date()%></font>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="录   入" id="submit1" name="submit1" onclick="return chk_form(this.form)">
      </td>
    </tr>
    <tr height=5>
    
     
        </form>

              
        
             

<!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          

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

tabid = trim(Request.form("tabid"))
if tabid=""then
tabid ="无"
end if

djdate=date()
djname=session("emid")

mnydepm=trim(f)

mnykm=trim(km)

mnynote=trim(Request.Form("mnynote"))
if mnynote=""then
mnynote="无"
end if
 


payway = trim(Request.form("payway"))



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
sqlc="SELECT * FROM cwys_ed WHERE depar='"& f &"' and ys_year='"&mnyyear&"' and fkmcode='"&trim(fkmcode1)&"'"                                                                            
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
                                                                           
sql7="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode1)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

rs_7.Open sql7,conn_1,3,3,1
'Response.Write(sql7)
if not rs_7.EOF then


q=cint(mnymonth)


if q="1" then
yxmy=rs_7("jan")
date2=mnyyear&"-"&"1"&"-"&"1"
date2=cdate(date2)
end if

if q="2" then
yxmy=rs_7("jan")+rs_7("feb")
date2=mnyyear&"-"&"2"&"-"&"1"
date2=cdate(date2)
end if

if q="3" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")
date2=mnyyear&"-"&"3"&"-"&"1"
date2=cdate(date2)
end if

if q="4" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")
date2=mnyyear&"-"&"4"&"-"&"1"
date2=cdate(date2)
end if

if q="5" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")
date2=mnyyear&"-"&"5"&"-"&"1"
date2=cdate(date2)
end if

if q="6" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")
date2=mnyyear&"-"&"6"&"-"&"1"
date2=cdate(date2)
end if


if q="7" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")
date2=mnyyear&"-"&"7"&"-"&"1"
date2=cdate(date2)
end if

if q="8" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")
date2=mnyyear&"-"&"8"&"-"&"1"
date2=cdate(date2)
end if

if q="9" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")
date2=mnyyear&"-"&"9"&"-"&"1"
date2=cdate(date2)
end if

if q="10" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")
date2=mnyyear&"-"&"10"&"-"&"1"
date2=cdate(date2)
end if

if q="11" then
yxmy=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")+rs_7("nov")
date2=mnyyear&"-"&"11"&"-"&"1"
date2=cdate(date2)
end if

if q="12" then
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
                                                                           
sql6="SELECT sum(price)as mnyused FROM cwys_infoin where mnykmcode='"&trim(fkmcode1)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' and mnytime<='"&date2&"' and ifhandin='是' and cz<>'删除' group by mnykmcode,mnydepm"
'Response.Write(sql6)
rs_6.Open sql6,conn_1,3,3,1
if rs_6.EOF then
mnyused=0
else

mnyused=rs_6("mnyused")
end if
%>



<%

'到用户所zai月份能用的钱的总和

set rs_77=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql77="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode1)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

rs_77.Open sql77,conn_1,3,3,1
'Response.Write(sql77)
if not rs_77.EOF then
p=cint(Month(date))

if p="1" then
yxmy1=rs_7("jan")
end if

if p="2" then
yxmy1=rs_7("jan")+rs_7("feb")
end if

if p="3" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")
end if

if p="4" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")
end if

if p="5" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")
end if

if p="6" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")
end if

if p="7" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")
end if

if p="8" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")
end if

if p="9" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")
end if

if p="10" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")
end if

if p="11" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")+rs_7("nov")
end if

if p="12" then
yxmy1=rs_7("jan")+rs_7("feb")+rs_7("mar")+rs_7("apr")+rs_7("may")+rs_7("jun")+rs_7("jul")+rs_7("aug")++rs_7("sep")+rs_7("oct")+rs_7("nov")+rs_7("dece")
end if



end if 

%>

<%
'到所zai月份用户已经使用过的钱

set rs_66=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql66="SELECT sum(price)as mnyused FROM cwys_infoin where mnykmcode='"&trim(fkmcode1)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' and mnytime<='"&date()&"' and ifhandin='是' and cz<>'删除' group by mnykmcode,mnydepm"
'Response.Write(sql6)
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
sqla="insert into cwys_infoin (passcode,passname,bxcode,bxname,djdate,djname,mnydepm,mnykm,mnynote,price,mnytime,payway,ifhandin,ifhx,mnyyear,mnykmcode,cz,tabid) values ('"&passcode&"','"&passname&"','"&bxcode&"','"&bxname&"','"&djdate&"','"&djname&"','"&mnydepm&"','"&+trim(km2)+"/"+trim(km)&"','"&mnynote&"',"&price&",'"&date1&"','"&payway&"','"&ifhandin&"','否','"&mnyyear&"','"&fkmcode1&"','录入','"&tabid&"')"                                                                            
'Response.Write(sqla)
rs_a.Open sqla,conn_1,3,3,1




%>

<table bgcolor=cccccc width=650>

<tr bgcolor=e3e3e3>
<td><font class="px12">所选年度</font></td>
<td><font class="px12">所选月份</font></td>
<td><font class="px12">允许额度</font></td>
<td><font class="px12">截止<%=mnymonth%>月份累计已经使用</font></td>
<td><font class="px12">剩余额度</font></td>
<td><font class="px12">录入结果</font></td>
</tr>
<tr bgcolor=white>

<td><font class="px12"><%=mnyyear%></font></td>
<td><font class="px12"><%=mnymonth%></font></td>
<td><font class="px12"><%=yxmy%>元</font></td>
<td><font class="px12"><%=mnyused+price%>元</font></td>
<td><font class="px12"><%=chae%>元</font></td>

<td><font class="px12">已经入库，已经提交财务</font></td>
</tr>

</table>


 <%else%>
 
 
<%'Response.Write("bunengcun")
ifhandin="否"
set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2                                                                            
sqlb="insert into cwys_infoin (passcode,passname,bxcode,bxname,djdate,djname,mnydepm,mnykm,mnynote,price,mnytime,payway,ifhandin,ifhx,mnyyear,mnykmcode,cz,tabid) values ('"&passcode&"','"&passname&"','"&bxcode&"','"&bxname&"','"&djdate&"','"&djname&"','"&mnydepm&"','"&+trim(km2)+"/"+trim(km)&"','"&mnynote&"',"&price&",'"&date1&"','"&payway&"','"&ifhandin&"','否','"&mnyyear&"','"&fkmcode1&"','录入','"&tabid&"')"                                                                            
'Response.Write(sqlb)
rs_b.Open sqlb,conn_1,3,3,1


%>
    

<table bgcolor=cccccc width=650>

<tr bgcolor=e3e3e3>
<td><font class="px12">所选年度</font></td>
<td><font class="px12">所选月份</font></td>

<td><font class="px12">允许额度</font></td>
<td><font class="px12">截止<%=mnymonth%>月份累计已经使用</font></td>
<td><font class="px12">剩余额度</font></td>
<td><font class="px12">目前需财务调整额度</font></td>
<td><font class="px12">录入结果</font></td>
</tr>
<tr bgcolor=white>

<td><font class="px12"><%=mnyyear%></font></td>
<td><font class="px12"><%=mnymonth%></font></td>

<td><font class="px12"><%=yxmy%>元</font></td>
<td><font class="px12"><%=mnyused%>元</font></td>

<td><font class="px12"><%=shengyu%>元</font></td>
<td><font class="px12" color="blue"><%=price-shengyu%>元</font></td>
<td><font class="px12" color="red">已经入库，没有提交财务!!!</font></td>

</tr>

</table>
 
 
  
 
 <%end if%>
 
 <%else%>
   
   <%'存
ifhandin="是"
'Response.Write("nengcun")
'Response.Write(price)



%>

<%

set rs_l=server.CreateObject("adodb.recordset")                                                                            
rs_l.CursorLocation=2                                                                            
sqll="insert into cwys_infoin (passcode,passname,bxcode,bxname,djdate,djname,mnydepm,mnykm,mnynote,price,mnytime,payway,ifhandin,ifhx,mnyyear,mnykmcode,cz,isover,tabid) values ('"&passcode&"','"&passname&"','"&bxcode&"','"&bxname&"','"&djdate&"','"&djname&"','"&mnydepm&"','"&+trim(km2)+"/"+trim(km)&"','"&mnynote&"',"&price&",'"&date1&"','"&payway&"','"&ifhandin&"','否','"&mnyyear&"','"&fkmcode1&"','录入','是','"&tabid&"')"                                                                            
'Response.Write(sqla)
rs_l.Open sqll,conn_1,3,3,1

%>

<%

set rs_o=server.CreateObject("adodb.recordset")                                                                            
rs_o.CursorLocation=2                                                                            
sqlo="update cwys_ed set isover='0' where depar='"& f &"' and ys_year='"&mnyyear&"' and fkmcode='"&trim(fkmcode1)&"'"                                                                            
'Response.Write(sqla)
rs_o.Open sqlo,conn_1,3,3,1


%>

<table bgcolor=cccccc width=650>

<tr bgcolor=e3e3e3>
<td><font class="px12">年度</font></td>
<td><font class="px12">月份</font></td>

<td><font class="px12">允许额度</font></td>
<td><font class="px12">截止<%=mnymonth%>月份累计已经使用</font></td>
<td><font class="px12">剩余额度</font></td>
<td><font class="px12">录入结果</font></td>
</tr>
<tr bgcolor=white>

<td><font class="px12"><%=mnyyear%></font></td>
<td><font class="px12"><%=mnymonth%></font></td>

<td><font class="px12"><%=yxmy%>元</font></td>
<td><font class="px12"><%=mnyused+price%>元</font></td>
<td><font class="px12">0元</font></td>
<td><font class="px12">已经入库，已经提交财务,此次提交由于财务允许超标</font></td>
</tr>

</table>
 
      
<%end if%>
<%end if%>
    <table border="0" cellspacing="0" cellpadding="0" width="600" align="left">
     <tr>
       <td width="600" height="20" align="left" bgcolor="#afc9e4">
       
        <iframe name="detail1" allowTransparency="true" src="ccs_infointable.asp?km=<%=km%>&km1=<%=trim(fkmcode1)%>&km2=<%=trim(km2)%>" width="800" height="210" align="center" frameborder="no">
        </iframe>
            </td>
      </tr>
</table>
<%end sub
%>





<!--显示剩余数据//-->
  <% 
                                                                        
Sub list_data()                                                                          

'获得数据


mnydepm=trim(f)

mnykm=trim(km)


mnyyear=year(date())

set rs_c=server.CreateObject("adodb.recordset")                                                                            
rs_c.CursorLocation=2                                                                            
sqlc="SELECT * FROM cwys_ed WHERE depar='"& f &"' and ys_year='"&mnyyear&"' and fkmcode='"&trim(fkmcode1)&"'"                                                                            
rs_c.Open sqlc,conn_1,3,3,1
if rs_c.EOF then%>
<table bgcolor=cccccc width=650>

<tr bgcolor=e3e3e3>

<td><font class="px12">没有此项科目的额度,您的帐单没有入库,请联系财务为此项科目分配额度</font></td>
</tr>

</table>
<%else

%>









<%

'到用户所zai月份能用的钱的总和

set rs_77=server.CreateObject("adodb.recordset")                                                                            
                                                                           
sql77="SELECT * FROM cwys_ed where fkmcode='"&trim(fkmcode1)&"' and depar='"&trim(f)&"' and ys_year='"&trim(mnyyear)&"'" 

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
                                                                           
sql66="SELECT sum(price)as mnyused FROM cwys_infoin where mnykmcode='"&trim(fkmcode1)&"' and mnydepm='"&trim(f)&"' and mnyyear='"& mnyyear &"' and mnytime<='"&date()&"' and ifhandin='是' and cz<>'删除' group by mnykmcode,mnydepm"
'Response.Write(sql66)
rs_66.Open sql66,conn_1,3,3,1
if rs_66.EOF  then
mnyused1=0
else
mnyused1=rs_66("mnyused")

end if

if isnull(rs_66("mnyused")) then
mnyused1=0
end if
%>











<table bgcolor=003399 width=650 border=0>

<tr bgcolor=b5cdff>
<td><font class="px12">年度</font></td>
<td><font class="px12">月份</font></td>
<td><font class="px12">科目</font></td>
<td><font class="px12">截止<%=month(date())%>月份费用总额</font></td>
<td><font class="px12">已经使用费用</font></td>
<td><font class="px12">剩余额度</font></td>

</tr>
<tr bgcolor=white>

<td><font class="px12"><%=mnyyear%></font></td>
<td><font class="px12"><%=month(date())%></font></td>
<td><font class="px12"><%=km2%>(<%=km%>)</font></td>
<td><font class="px12"><%=yxmy1%></font></td>
<td><font class="px12"><%=mnyused1%></font></td>
<td><font class="px12"><%=yxmy1-mnyused1%></font></td>
</tr>

</table>


 
<%
end if
end sub
%>




<%'主过程
                                                                                                                        
Select case Request.QueryString("todo")                                                                         
case ""
list_data()
         
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
