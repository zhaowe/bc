<% 
 
if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"ccs_gsgly")
   if VerifyOk=false then
      session("errorNo")="000002"
      Response.Redirect "../sorry/sorry.asp"
   end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
%> 
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
function reload()
{

window.opener.document.location.reload();
//window

}
//-->
</script>
</head>
<html>
<body onload="javascript:reload()">
<%
sn=trim(Request.QueryString("sn"))
depar=trim(Request.form("depar"))
nian=trim(Request.form("nian"))
kemu=trim(Request.form("kemu"))
niandu=trim(Request.form("niandu"))
jan=trim(Request.form("jan"))
feb=trim(Request.form("feb"))
mar=trim(Request.form("mar"))
apr=trim(Request.form("apr"))
may=trim(Request.form("may"))
jun=trim(Request.form("jun"))
jul=trim(Request.form("jul"))
aug=trim(Request.form("aug"))
sep=trim(Request.form("sep"))
shiyiy=trim(Request.form("oct"))
nov=trim(Request.form("nov"))
dece=trim(Request.form("dece"))
isover=trim(Request.form("isover"))
beizhu=trim(Request.form("beizhu"))
%>
<%
dd=abs(cdbl(niandu)-cdbl(jan)-cdbl(feb)-cdbl(mar)-cdbl(apr)-cdbl(may)-cdbl(jun)-cdbl(jul)-cdbl(aug)-cdbl(sep)-cdbl(shiyiy)-cdbl(nov)-cdbl(dece))
''if clng(niandu)<>clng(jan)+clng(feb)+clng(mar)+clng(apr)+clng(may)+clng(jun)+clng(jul)+clng(aug)+clng(sep)+clng(shiyiy)+clng(nov)+clng(dece) then
if dd>0.001  then
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>����Ԥ���ȵ��ۼ�ֵ������ܶһ�£��뷵�ؼ�顣</font></p>
	<input type="button" name="button" value="����" onclick="JavaScript:history.go(-1)">
</td>
</tr>
</table>
<%else%>
<%
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("OledbStr") 
      
Set obj=server.CreateObject ("ADODB.Recordset")
obj.LockType=3
obj.CursorType=3
set obj.activeConnection=objConn
sql="select * FROM cwys_ed where sn='"&sn&"'"
obj.Source=sql
obj.Open
%>
<%'������д����־
descr=descr+"�޸�"+depar+nian+"��"+kemu+"��"
if trim(obj("niandu"))<>niandu then 
descr=descr+"��ȶ���ɣ�"+trim(obj("niandu"))+"�޸�Ϊ��"+niandu+"��"
end if

if trim(obj("jan"))<>jan then
descr=descr+"һ�·ݶ���ɣ�"+trim(obj("jan"))+"�޸�Ϊ��"+jan+"��"
end if

if trim(obj("feb"))<>feb then
descr=descr+"���·ݶ���ɣ�"+trim(obj("feb"))+"�޸�Ϊ��"+feb+"��"
end if

if trim(obj("mar"))<>mar then
descr=descr+"���·ݶ���ɣ�"+trim(obj("mar"))+"�޸�Ϊ��"+mar+"��"
end if

if trim(obj("apr"))<>apr then 
descr=descr+"���·ݶ���ɣ�"+trim(obj("apr"))+"�޸�Ϊ��"+apr+"��"
end if

if trim(obj("may"))<>may then
descr=descr+"���·ݶ���ɣ�"+trim(obj("may"))+"�޸�Ϊ��"+may+"��"
end if

if trim(obj("jun"))<>jun then
descr=descr+"���·ݶ���ɣ�"+trim(obj("jun"))+"�޸�Ϊ��"+jun+"��"
end if

if trim(obj("jul"))<>jul then
descr=descr+"���·ݶ���ɣ�"+trim(obj("jul"))+"�޸�Ϊ��"+jul+"��"
end if

if trim(obj("aug"))<>aug then
descr=descr+"���·ݶ���ɣ�"+trim(obj("aug"))+"�޸�Ϊ��"+aug+"��"
end if

if trim(obj("sep"))<>sep then
descr=descr+"���·ݶ���ɣ�"+trim(obj("sep"))+"�޸�Ϊ��"+sep+"��"
end if

if trim(obj("oct"))<>shiyiy then
descr=descr+"ʮ�·ݶ���ɣ�"+trim(obj("oct"))+"�޸�Ϊ��"+shiyiy+"��"
end if

if trim(obj("nov"))<>nov then
descr=descr+"ʮһ�·ݶ���ɣ�"+trim(obj("nov"))+"�޸�Ϊ��"+nov+"��"
end if

if trim(obj("dece"))<>dece then
descr=descr+"ʮ���·ݶ���ɣ�"+trim(obj("dece"))+"�޸�Ϊ��"+dece+"��"
end if

if trim(obj("isover"))<>isover then
descr=descr+"�Ƿ�ɳ�֧�ɣ�"+trim(obj("isover"))+"�޸�Ϊ��"+isover+"��"
end if

if beizhu<>"" then
descr=descr+"�޸�ԭ��Ϊ��"+beizhu+"��"
else
descr=descr+"δ��д�޸�ԭ��"
end if

set rs_b=server.CreateObject("adodb.recordset")                                                                            
rs_b.CursorLocation=2  
sqlb="insert into cwys_log (operation,descr,type,operatetime,operator) values ('�޸�','"&descr&"','Ԥ��','"&now&"','"&session("emid")&"')" 
rs_b.Open sqlb,objConn,3,3,1 
'rs_b.Close

%>

<%
obj("niandu")=niandu
obj("jan")=jan
obj("feb")=feb
obj("mar")=mar
obj("apr")=apr
obj("may")=may
obj("jun")=jun
obj("jul")=jul
obj("aug")=aug
obj("sep")=sep
obj("oct")=shiyiy
obj("nov")=nov
obj("dece")=dece
obj("isover")=isover                                                                                                                      
obj.Update                                                                          
obj.close  
%>
<table width=500>
<tr align="center" width=500>
<td align="center" width=500>
	<font class=px14 color=blue>�޸ĳɹ���</font></p>
	<input type="button" name="button" value="ȷ��" onclick="JavaScript:window.close()">
</td>
</tr>
</table>
<%end if%>
</body>
</html>
