<%@ Language=VBScript %>

<% 
 
'if trim(session("UID"))<>"" then
 '  dim objD
 '  set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
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

  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn%>
<html>
<head>
<title>���ڹ�˾Ԥ�����ϵͳ</title>
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
.contentindex{font-family: "����";FONT-SIZE:9pt; LINE-HEIGHT:11pt;}
.enter {COLOR: #FFAF02; FONT-FAMILY: "����", "Arial", "Times New Roman"; FONT-SIZE: 11pt; TEXT-DECORATION: none ;font-weight: bold}
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
                <td><a href><img src="images/lizztp4.gif" width="138" height="28" border="0"></a></td>
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
                     <td><a href="ccs_xtwf.asp"><font class="px14" color="#FFFFFF">��Ŀ¼��</font></a></td>
                      </tr>
                      <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="ccs_xtwf_xg.asp"><font class="px14" color="#FFFFFF">��Ŀ�޸ļ�ɾ��</font></a></td>
                       </tr>
                        <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="ccs_xtwf_kh.asp"><font class="px14" color="#FFFFFF">���˿�Ŀ¼��</font></a></td>
                       </tr>
                        <tr>
                      <td><img src="images/a.gif" WIDTH="28" HEIGHT="28"></td>
                      <td><a href="ccs_khxtwf_xg.asp"><font class="px14" color="#FFFFFF">���˿�Ŀ�޸ļ�ɾ��</font></a></td>
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
    <table align="center" style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="30" cellSpacing="0" cellPadding="0" width="775"  border="0">
  <tbody>
      
   <tr>
    <td align="right" >&nbsp;</td></tr>
   <tr>
    <td align="center" ><font color=black class=px12>��Ŀǰ������ǿ�Ŀ�޸ļ�ɾ������</td></tr>
    <tr>
    <td align="right" >&nbsp;</td></tr>
   </tbody></table>
   <% objrst.Source = "select * from shenzhencwys_dep  where  dep='"& session("dep")&"' order by kem,sx "
   objRst.Open 
   if objrst.EOF and objrst.BOF   then %>    
     <P> 
     <P align=center><font color=black class=px12><STRONG>û����Ҫ��<%=km%>����Ϣ
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
          <td style="PADDING-LEFT: 10px; PADDING-TOP: 2px" bgColor="#4983a0" height="20"><font style="COLOR: #ffffff; FONT-FAMILY: Webdings">4</font><font color="#ffffff"><b></b></font></td></tr>
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#ecf7fd" height="20">
                <td align="middle"  ><font color=black class=px12>����</td>
                <td align="middle" ><font color=black class=px12>��Ŀ</td>
                <td align="middle"  ><font color=black class=px12>�ӿ�Ŀ</td>
                <td align="middle"  ><font color=black class=px12>˳��</td>
                <td align="middle"  ><font color=black class=px12>����</td>
                </tr>
                <%
                i=1
                do while not objrst.EOF and i<=21 %>
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle"  ><a href="ccs_lu_xg.asp?q=<%=objrst("q")%>&km=<%=km%>"><font color=blue class=px12>�޸�</font></a></td>
                <td align="middle" ><font color=black class=px12><%=objrst("kem")%></td>
                <td align="middle" ><font color=black class=px12><%=objrst("fkem")%></td>
                <td align="middle"  ><font color=black class=px12><%=objrst("sx")%></td>
                <td align="middle"  ><a href="ccs_lu_del.asp?q=<%=objrst("q")%>&km=<%=km%>"><font color=red class=px12>ɾ��</font></a></td>
                </tr>
                <%i=i+1
                objrst.MoveNext 
                loop%> 
          
             </tbody></table></td></tr></tbody></table>
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
             <br>



              <br> <br> <br>
               <br> <br> <br>
      <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><font class="px12">Copyright 2006, �й��Ϸ��������ڹ�˾��Ϣ���̲�</font></td>
        </tr>
      </table>
      <p>��</p>
    </td>
  </tr>
</table>
</body>
</html>


