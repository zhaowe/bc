

<%@ Language=VBScript %>


  
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
                <td><a href="kmgl.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images2','','images/kemu1.gif',1)"><img src="images/kemu.gif" width="120" height="24" border="0" name="images2"></a></td>
                <td><a href="ysgl.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images3','','images/yusuan1.gif',1)"><img src="images/yusuan.gif" width="120" height="24" border="0" name="images3"></a></td>
                <td><a href="cx/ccs_gscxy_index.asp" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapimages('images4','','images/gltj2.jpg',1)"><img src="images/gltj1.jpg" width="120" height="24" border="0" name="images4"></a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
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
    
    <%sub list_data()%> 
    <form method="post" action="ccs_gkbm_add.asp?todo=02" id=form1 name=form1>   
     <table style="BORDER-RIGHT: #4e4c71 1px solid; BORDER-TOP: #4e4c71 1px solid; BORDER-LEFT: #4e4c71 1px solid; BORDER-BOTTOM: #4e4c71 1px solid" height="92%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td vAlign="top" width="100%">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
              
            


<tr width=100%>
     <td  align="left">
     <font class="px12" color="black">&nbsp;&nbsp;子科目代码：</font>
     <input id="hbh" type="text" name="fkmcode" size="10" ><font class="px12" color="black">  </font>  
    
     <font class="px12" color="black">&nbsp;&nbsp;部门：</font>
     
      <select name="bm"> 
      
   
      <% 
          Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
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
     
   
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="录   入" id="submit1" name="submit1">
      </td>
    </tr>
    <tr height=5>
    
     
        </form>

     
     <% 
     
      Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
     objrst.Source = "select * from cwys_gkbm order by department,fkmcode"
   objRst.Open
   'Response.Write(objrst.source) 
   if objrst.EOF and objrst.BOF   then %>    
    <font color=black class=px12><STRONG>没有你要的<%=km%>的信息
   <%else%>
   
<table align="left"  cellSpacing="0" cellPadding="0" width="750" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="100%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="561" border="0">
        <tbody>

        <tr>
          <td vAlign="top" width="564">
            <table cellSpacing="1" cellPadding="0" width="750">
              <tbody>
               <tr bgColor="#9CD7F5" height="20" width="750">
               
                <td align="middle" width="64"  ><font color=black class=px12>记录号</td>
                <td align="middle" width="64"  ><font color=black class=px12>部门</td>
                <td align="middle" width="64"  ><font color=black class=px12>子科目代码</td>
                <td align="middle" width="63"  ><font color=black class=px12>操作</td>
                
                </tr>
                <%
               
                do while not objrst.EOF%>
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("sn")%></td>
          
                <td align="middle" width="64"  ><font color=black class=px12><%=objrst("department")%></td>
                    <td align="middle" width="64"  ><font color=black class=px12><%=objrst("fkmcode")%></td>
                <td align="middle" width="63"  ><a href="ccs_gkbm_add.asp?todo=01&sn=<%=trim(objrst("sn"))%>" onclick="return check()"><font color=red class=px12>删除</font></a></td>
             
                </tr>
                <%
                objrst.MoveNext 
                loop%> 
          
             </tbody></table></td></tr></tbody></table>
<%end if%>

</table>   
<%end sub%>      
        
             

<!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          

'获得数据
fkmcode = trim(Request.form("fkmcode"))


depar = trim(Request.form("bm"))



set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")
    
 set rs1=server.CreateObject("adodb.recordset") 
 rs1.CursorLocation=2 
sql="select count(*) as num from cwys_gkbm where department='"&depar&"'and fkmcode='"&fkmcode&"'"
rs1.Open sql,conn_1,3,3,1

if rs1("num")=0 then




  set rs1=server.CreateObject("adodb.recordset") 
 rs1.CursorLocation=2                                                                           
sqll="insert into cwys_gkbm (department,fkmcode) values ('"&depar&"','"&fkmcode&"')"                                                                            
'Response.Write(sqla)
rs1.Open sqll,conn_1,3,3,1

else

%>
        <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><font class="px12">本条记录已经存在，请重新输入</font></td>
        </tr>
      </table>   

<%
end if
end sub
%>

<%
sub del_data()
sn=Request.QueryString("sn")
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")
    set rs_l=server.CreateObject("adodb.recordset")                                                                            
rs_l.CursorLocation=2                                                                            
sqll="delete cwys_gkbm where sn='"&sn&"'"                                                                            
'Response.Write(sqla)
rs_l.Open sqll,conn_1,3,3,1
end sub
%>


<%'主过程
                                                                                                                        
Select case Request.QueryString("todo")                                                                         
case ""
list_data()
       case "01"
       del_data() 
       list_data() 
       case"02" 
       save_data() 
       list_data() 
                                                      
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
