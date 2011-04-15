<%@ Language=VBScript %>
<%

   m_errorNo=session("errorNo")

   'response.write(Application("OledbStr"))
   Set Conn1 = Server.CreateObject("ADODB.Connection")
   Conn1.Open Application("OledbStr") 

   
   
   Set Rs1=server.CreateObject ("ADODB.Recordset")
   Rs1.LockType=3
   Rs1.CursorType=3
   set Rs1.activeConnection=Conn1
      
   Rs1.Source="select * from error where Errornumber='" &trim(m_errorNo)& "'"
   'Response.Write Rs1.Source
   Rs1.Open  


   if Rs1.eof and Rs1.bof then
      ErrorName="对不起，你的工作号或者密码使用不正确！"
	  Solution=""
     else
      ErrorName=Rs1(1)
	  Solution=Rs1(2)
   end if  
                                                                                                   
%>

<script language="JavaScript">
<!--
function back_onclick() { window.history.back(); return true; }

//-->
</script>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title></title>
</head>
<body>
<style type="text/css">
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
<div id="Layer1" style="HEIGHT: 41px; LEFT: 1px; POSITION: absolute; TOP: 1px; WIDTH: 200px; Z-INDEX: 1"><img src="images/index_new_2_r1_c1.gif" WIDTH="800" HEIGHT="77"></div>
<div id="Layer4" style="HEIGHT: 41px; LEFT: 600px; POSITION: absolute; TOP: 290px; WIDTH: 200px; Z-INDEX: 2">
	<img border="0" src="images/dl-1.gif" WIDTH="66" HEIGHT="40" LANGUAGE=javascript onclick="return back_onclick()">
	
</div>

<div id="Layer2" style="BACKGROUND-IMAGE: url(images/index_new_2_r8_c1.gif); HEIGHT: 128px; LEFT: 1px; POSITION: absolute; TOP: 274px; WIDTH: 800px; Z-INDEX: 1; layer-background-image: url(images/index_new_2_r8_c1.gif)">
<br>
<br>

 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="white" class="px14">如果您有任何问题或建议，请联络<a href="mailto:lizz@cs-air.com"><font color="white" class="px14">管理员</font></a></font>
                        
</div>
 
 <div id="Layer3" style="HEIGHT: 41px; LEFT: 50px; POSITION: absolute; TOP: 80px; WIDTH: 200px; Z-INDEX: 1"><img src="images/index_new_2_r2_c1.gif" WIDTH="171" HEIGHT="178">
 
 </div>
<div id="Layer5" style="HEIGHT: 161px; LEFT: 219px; POSITION: absolute; TOP: 98px; WIDTH: 500px; Z-INDEX: 1">
 <table border="0" width="100%" cellspacing="4" cellpadding="2">
      <tr>
         <td bgcolor="#ffffff">
               <table border="0" bgcolor="#ffffff">
               </table>

                <p align="center"><font color="black" class="px16"><b> 以下是您出错信息的提示</b></font></p>

                <table align="center" border="1" width="100%" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="100" align="center"><font color="red" class="px14">出错原因：</font></td>
                    <td><font color="red" class="px14"><% =errorName%></font></td>
                  </tr>
 </table>

 </div>
<%
  Conn1.Close
  set rs1 = nothing
  set Conn1 = nothing
%> 

</body>
</html>
