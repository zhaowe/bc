
<%@ Language=VBScript %>


  
<html>
<head>
<title>公司预算管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">


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
if rs_k.EOF then
Response.Write("本年度您所在部门的各项额度尚未录入，请与财务部门联系！")
else


km=Request.QueryString ("km")
mnykm=Request.QueryString("mnykm")
if km="" then
if mnykm="" then
'km=trim(rs_k("fkmshuom"))
km=session("km")
else
km=mnykm
end if
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
 
 km2=Request.QueryString ("km2")
 if km2="" then
 km2=session("km2")
 end if
  km1=Request.QueryString ("km1")
  if km1="" then
  km1=session("km1")
  end if
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
.news{FONT-SIZE:10pt; LINE-HEIGHT:14pt; color:#000000:#000000; }
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
function xx(j,k,l,m,n,o,p) { //v3.0
   //s=depar.value ;
   //y=nian.value ;
   surl="ccs_bm_ser.asp?mnykm="+j+"&year1="+k+"&month1="+l+"&year2="+m+"&month2="+n+"&km2="+o+"&km1="+p;
   window.location.href (surl);
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

function check()
{
if (confirm("你确定要删除吗？")==false)
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
//function dkxg()
//{

//}


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

                <td><a href="ccs_bm_ser.asp?km=<%=km%>&km1=<%=km1%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images1','','images/ysglinfoin1.jpg',1)"><img src="images/ysglinfoin2.jpg" width="120" height="25" border="0" name="images1"></a></td>
              
                <td><a href="ccs_input_index.asp?km=<%=km%>&km1=<%=km1%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/bt_02_off.gif',1)"><img src="images/bt_02_on.gif" width="120" height="25" border="0" name="images2"></a></td>
                <td><a href="ccs_bmgl_zttj.asp?km=<%=km%>&km1=<%=km1%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/gltj2.jpg',1)"><img src="images/gltj1.jpg" width="120" height="24" border="0" name="images4"></a></td>
                <td><a href="ccs_bm_rzcx.asp?km=<%=km%>&km1=<%=km1%>&km2=<%=trim(km2)%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images3','','images/bmrz2.jpg',1)"><img src="images/bmrz1.jpg" width="120" height="24" border="0" name="images3"></a></td>
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
                <td><a href><img src="images/bmysglsm1.gif" width="138" height="28" border="0"></a></td>
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
                  objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '"&f&"' order by kmcode,kmshuom" 
    
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
                   objrst.Source ="select distinct kmshuom,kmcode,fkmshuom,fkmcode from cwys_km where depar = '"&f&"' and kmshuom='"&trim(objrst1("kmshuom"))&"' order by kmcode,kmshuom,fkmcode,fkmshuom" 
                       'Response.Write(objrst.Source)
                     
                       objrst.Open 
                       while not objrst.EOF
                   
                   %>

                  <TR>
                    <TD style="PADDING-LEFT: 0px" height=23> <A title=资料册 
                        href="ccs_bm_ser.asp?km=<%=trim(objrst("fkmshuom"))%>&km1=<%=trim(objrst("fkmcode"))%>&km2=<%=trim(objrst("kmshuom"))%>" 
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
    <td width="650" valign="top"> 
     <table style="BORDER-RIGHT: #4e4c71 1px solid; BORDER-TOP: #4e4c71 1px solid; BORDER-LEFT: #4e4c71 1px solid; BORDER-BOTTOM: #4e4c71 1px solid" height="92%" cellSpacing="0" cellPadding="0" width="723" border="0">
        <tbody>
        <tr>
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4e4c71" height="20" width="711"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff" class="px14">&nbsp;科目管理</font> </td></tr>
        <tr>
          <td vAlign="top" width="721">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
              
              <tr bgColor="#d8d9de" height="20">
    <td align="right" >　</td></tr>
    
   <tr>
    <td align="center" ><font class="px12" color="black">您目前进入的是<font color=red><%=km2%>(<%=km%>)</font> <font color=blue>管理</font>界面</td></tr>
    <tr>
    <td align="right" >　</td></tr>



    <tr>    <td align="left">
       <font class="px12" color="black">费用科目：</font>
     <input type="hidden" name="km2" value=<%=km2%>><input type="hidden" name="km1" value=<%=km1%>><input type="text" name="mnykm" size="20" readonly value=<%=km%>>
<%set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")                                                                           
set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2                                                                            
sql2="SELECT fkmcode FROM cwys_km WHERE fkmshuom='"& km &"' and depar='"&f&"' and nian='"&year(date())&"'"                                                                            
rs_2.Open sql2,conn_1,3,3,1
if not rs_2.EOF then
fkmcode1=rs_2("fkmcode")
'Response.Write(fkmcode1)
else
fkmcode1="无"
end if
%>
<font class="px12"><%=fkmcode1%></font>    
  
    <font class="px12" color="black">&nbsp;费用期间：</font>
        <select size="1" name="year1">
         <OPTION  value="<%=cstr(year(date)+1)%>"><font color=black class="px14"><%=cstr(year(date)+1)%></OPTION>
             <OPTION selected value="<%=cstr(year(date))%>"><font color=black class="px14"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><font color=black class="px14"><%=cstr(year(date)-1)%></OPTION>
                       <OPTION  value="<%=cstr(year(date)-2)%>"><font color=black class="px14"><%=cstr(year(date)-2)%></OPTION>

            </select>
            <font color=black class="px14">年</font>
        <select name="month1" size="1"> 
        <option selected value="<%=month(date)%>"><%=month(date)%>
        <%dd=13
        
          cha=dd-1
          while cha>0 
    %>
        <option value="<%=cha%>"><%=cha%></option>
        
        <%
        cha=cha-1
        wend%></select>
        
      <font class="px12" color="black">月至</font>        
      <select size="1" name="year2">
         <OPTION  value="<%=cstr(year(date)+1)%>"><font color=black class="px14"><%=cstr(year(date)+1)%></OPTION>
             <OPTION selected value="<%=cstr(year(date))%>"><font color=black class="px14"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><font color=black class="px14"><%=cstr(year(date)-1)%></OPTION>
                       <OPTION  value="<%=cstr(year(date)-2)%>"><font color=black class="px14"><%=cstr(year(date)-2)%></OPTION>

            </select>
            <font color=black class="px14">年</font>
        <select name="month2" size="1"> 
        <option selected value="<%=month(date)%>"><%=month(date)%>
        <%dd=13
        
          cha=dd-1
          while cha>0 
    %>
        <option value="<%=cha%>"><%=cha%></option>
        
        <%
        cha=cha-1
        wend%></select>
        
      <font class="px12" color="black">月
     

<input type="submit" value="查询" id="submit1" name="submit1" onclick="javascript:xx(mnykm.value,year1.value,month1.value,year2.value,month2.value,km2.value,km1.value)">
        
      </td>
 






        
    <tr height=5>
    </table>          
           
             

<!--保存数据//-->
  <% 
                                                                        
'Sub save_data()                                                                          

'获得数据


'mnykm=trim(km)

'mnyyear1 = trim(Request.form("year1"))
'mnymonth1 = trim(Request.Form("month1"))
'mnyyear2 = trim(Request.form("year2"))
'mnymonth2 = trim(Request.Form("month2"))

mnykm=Request.QueryString("mnykm")
mnyyear1=Request.QueryString("year1")
mnyyear2=Request.QueryString("year2")
mnymonth1=Request.QueryString("month1")
mnymonth2=Request.QueryString("month2")
if mnymonth1="" then
date1=""
date2=""
mnyyear1=""
mnyyear2=""
mnymonth1=""
mnymonth2=""
else
 date1=mnyyear1&"-"&mnymonth1&"-"&"1"
 date2=mnyyear2&"-"&mnymonth2&"-"&"1"
end if
%>



<%

 
set rs_7=server.CreateObject("adodb.recordset")                                                                            
 
cz=Request.QueryString("cz") 
 if cz="删除" then
sql7=session("sql7")
else
 
 
                                                                           
sql7="SELECT * FROM cwys_infoin where mnykmcode='"&km1&"' and mnytime>='"&date1&"' and mnytime<='"&date2&"' and cz<>'删除' and mnydepm='"&f&"' order by ifhx,mnytime desc" 
end if
rs_7.Open sql7,conn_1,3,3,1
'Response.write(sql7)
session("sql7")=sql7
if rs_7.EOF then%>
<br>
<table><font class="px12">请输入查询条件</font></table>
<%else
%>
<br>
<table bgcolor="white" width="764" cellSpacing="1" cellPadding="0">

       <tr bgColor="#7dadc4" height="20">
                <td align="middle"  ><font color=black class=px12><b>操作</b></td>
                <td align="middle"  ><font color=black class=px12><b>费用期间</b></td>
                <td align="middle"  ><font color=black class=px12><b>付款方式</b></td>
                <td align="middle"  ><font color=black class=px12><b>子科目</b></td>
                <td align="middle"  ><font color=black class=px12><b>金额</b></td>
                <td align="middle"  ><font color=black class=px12><b>是否提交</b></td>
                <td align="middle"  ><font color=black class=px12><b>简要说明</b></td>
                <td align="middle"  ><font color=black class=px12><b>报销人</b></td>
                <td align="middle"  ><font color=black class=px12><b>登记人</b></td>
                <td align="middle"  ><font color=black class=px12><b>经办人</b></td>
                <td align="middle"  ><font color=blue class=px12><b>是否核销</b></td>
                <td align="middle"  ><font color=black class=px12><b>记录号</b></td>
                <td align="middle"  ><font color=black class=px12><b>帐单号</b></td>
                <td align="middle"  ><font color=black class=px12><b>操作</b></td>
          
              </tr>
               
               <% do while not rs_7.EOF%>
                
              <tr bgcolor="ecf7fd">
              <%if trim(rs_7("ifhx"))="否" then
              %> 
    
               <td align="middle"  ><a href="#" onclick="JavaScript:window.open('ccs_lu_xg1.asp?q=<%=rs_7("record_id")%>&date1=<%=date1%>&date2=<%=date2%>&km1=<%=km1%>','hh','width=550,left=200,top=10,height=242');"><font color=blue class=px12>修改/提交</font></a></td>
                    
                <td align="middle" ><font color=black class=px12><%=rs_7("mnyyear")%>年<%=month(rs_7("mnytime"))%>月</td>
                <td align="middle" ><font color=black class=px12><%=rs_7("payway")%></td>
                <td align="middle" ><font color=black class=px12><%=rs_7("mnykm")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("price")%></td>
                <%if trim(rs_7("ifhandin"))="否" then%>
                <td align="middle"  ><font color="#ff9900" class=px12><%=rs_7("ifhandin")%></td>
                <%else%>
                <td align="middle"  ><font color=black class=px12><%=rs_7("ifhandin")%></td>
                <%end if%>
                <td align="middle"  ><font color=black class=px12><%=rs_7("mnynote")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("bxname")%></td>

                <td align="middle"  ><font color=black class=px12><%=rs_7("djname")%></td>
                 <td align="middle"  ><font color=black class=px12><%=rs_7("passname")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("ifhx")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("record_id")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("tabid")%></td>
                <td align="middle"  ><a href="ccs_lu_del1.asp?q=<%=rs_7("record_id")%>&km=<%=km%>&km1=<%=km1%>&km2=<%=trim(km2)%>" onclick="return check()"><font color=red class=px12>删除</font></a></td>
      </tr>
      
      
      <%else%>
      
       <tr bgcolor="dbdbdb">          
                 
                 <td align="middle"  ><font color=666666 class=px12>修改</font></td>
                 <td align="middle" ><font color=black class=px12><%=rs_7("mnyyear")%>年<%=month(rs_7("mnytime"))%>月</td>
                <td align="middle" ><font color=black class=px12><%=rs_7("payway")%></td>
                <td align="middle" ><font color=black class=px12><%=rs_7("mnykm")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("price")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("ifhandin")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("mnynote")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("bxname")%></td>

                <td align="middle"  ><font color=black class=px12><%=rs_7("djname")%></td>

                <td align="middle"  ><font color=black class=px12><%=rs_7("passname")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("ifhx")%></td>
                <td align="middle"  ><font color=black class=px12><%=rs_7("record_id")%></td>
           <td align="middle"  ><font color=black class=px12><%=rs_7("tabid")%></td>
                <td align="middle"  ><font color=666666 class=px12>删除</font></td>
      
      
                </tr>
               
          <%end if%>      
                
                
                <%
                rs_7.MoveNext 
                loop
                
                end if%>
                
                
            
                
                <%'Rs_7.close                                                                            
                'set Rs_7=nothing%> 


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