
<%
function V2connect

Set V2connect = Server.CreateObject("ADODB.Connection")
'V2connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/newsv2.mdb")
V2connect.Open Application("OledbStr") 

end function



   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn





Set objRst=server.CreateObject ("ADODB.Recordset")
objRst.LockType=3
objRst.CursorType=3
set objRst.activeConnection=V2connect



'objRst.Source="select no,name from szairlineuser where logid='"& trim(session("LoginID")) &"'"
'objRst.Open 

'session("emid")=objrst("no") 
'objrst.close

session("emid")=196478
logid=trim(session("emid"))
sqlins="insert into xjxscore([no],[begin],[end],score) values('"& logid &"',getdate(),getdate(),0) "

objrst.Source =sqlins
objrst.open 



%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ƚ��Խ���</title>
<style type="text/css">
<!--
body,table {font-family: "����", "Arial", "Times New Roman";font-size: 10.5pt}
A:link,A:visited{color: yellow;TEXT-DECORATION: none;font-family: "����"}
A:hover		{color: red;   TEXT-DECORATION: none;font-family: "����"}
A.menu:link,A.menu:visited{color: yellow;TEXT-DECORATION: none;font-size: 10.5pt; font-family: "����"}
A.menu:hover	{color: red;   TEXT-DECORATION: none;font-size: 10.5pt; font-family: "����"}
A.blue:link,A.blue:visited{color: blue;TEXT-DECORATION: none;font-size: 10.5pt; font-family: "����"}
A.blue:hover	{color: red;   TEXT-DECORATION: none;font-size: 10.5pt; font-family: "����"}
.f9	{font-family: "����", "Arial", "Times New Roman";font-size: 9pt}
.f10	{font-family: "����", "Arial", "Times New Roman";font-size: 10.5pt}
.f12	{font-family: "����", "Arial", "Times New Roman";font-size: 12pt}
.f9y,.nav{font-family: "����", "Arial", "Times New Roman";font-size: 9pt;color: yellow}
-->
</style>
<script language=JavaScript>
<!--
var timerID=null
function showtime(seed){
if (seed>=0 && seed<=60 ){
seed++;
mod=seed%3600;
hours=(seed-mod)/3600
seconds=mod%60
minutes=(seed-3600*hours-seconds)/60;
var timeValue=""+((hours < 10)?"0":"")+hours
timeValue+=((minutes < 10)?":0":":")+minutes
timeValue+=((seconds < 10)?":0":":")+seconds
document.clock.face.value=timeValue;
timerID=timeValue;
var cmd="showtime("+seed+")";
timeID=window.setTimeout(cmd,1000) }

else
{
window.alert ("ʱ�䵽!") ;
checkscore(1);

}


}

//-->
</script>
</head>
<body topmargin="0" leftmargin="0">
<SCRIPT language=JavaScript src="float.js"></SCRIPT>
<div id="floater" style="position:absolute; width:84px; height:41px; z-index:1; left: 679px; top: 120px; visibility: visible">
  <form name=clock class=t9>
    <font color="#0033FF" class=f9>����ʱ��</font><BR>
    <input name=face style="font-size: 9pt;color:blue;border:0" size=10>
  </form>
</div>
<SCRIPT language=JavaScript>
var jiaojuan=0;
var XPos;
var YPos;
var ShowTime=0;
var isNetscape = navigator.appName=="Netscape";
var res = new Array;
var ans = new Array;
<% 
	dim NEWSconn, rs, rs2, sql
	Set NEWSconn=V2connect

	Randomize
	
	NEWSconn.Execute(" UPDATE [�Ƚ��Կ���] SET [�����] = cast((" & Rnd(1) & "*ID*1000) as int) % 431 ")

	sql="SELECT TOP 35 * from �Ƚ��Կ��� where ����=" & 1 & " ORDER BY ����� DESC , ID DESC;"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, NEWSconn, 3, 3
	dim i
	i = 1
	while not rs.eof
		response.write "ans[" & i & "]=" &  Asc(rs("��"))-64 & ";"
		rs.movenext
		i = i + 1
	wend

	sql="SELECT TOP 15 * from �Ƚ��Կ��� where ����=" & 2 & " ORDER BY ����� DESC , ID DESC;"
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.Open sql, NEWSconn, 3, 3
	while not rs2.eof 
		if rs2("��") = "T" then 
			response.write "ans[" & i & "]=1;"
		else
			response.write "ans[" & i & "]=2;"
		end if
		rs2.movenext
		i = i + 1
	wend
%>

function MoveHandler(e)
{
 XPos = e.pageX;
 YPos = e.pageY;
 return true;
}
// just save mouse position for animate() to use
function MoveHandlerIE() {
 XPos = window.event.x + document.body.scrollLeft;
 YPos = window.event.y + document.body.scrollTop;
}
if (isNetscape) {
 document.captureEvents(Event.MOUSEMOVE);
 document.onMouseMove = MoveHandler;
} else {
 document.onmousemove = MoveHandlerIE;
}
function record(question, answer) {
	if (jiaojuan==0)
		res[question] = answer;
	else{
		alert('���Ѿ������ˣ��������޸ģ��������ʦ��¼������');
		return false;
	}
};
//����ܷ�
function checkscore(num)
{if (jiaojuan==0) {
var score=0;
jiaojuan=1;
var sa;
for(var i=num; i < num+50; i++){ if (res[i]==ans[i]) score=score+2};
if (score>90) {sa=score+"�֣��ɼ����죡";}
else {if (score>75) {sa=score+"�֣��ɼ����á�";}
   else {if (score>60) {sa=score+"�֣��ɼ�һ�㡣";}
	else {if (score>40) {sa=score+"�֣��ɼ��д��Ľ���";}
		    else {sa=score+ "�֣�����ù�...!!";};
		    }
	}
};
document.all.se.value=score+"�֡���ʱ��"+timerID;
document.all.fen.value=score;
alert('�����'+sa);
cf.submit()
}
else
alert('���Ѿ����˾��Ҽ࿼��ʦ�ķְɡ�');
}
</SCRIPT>
<table border="0" width="777" cellspacing="0" cellpadding="0">
  <tr>
    <center>
      <td width="20%" bgcolor="#2163FF" valign="top" align="left" class=f9y rowspan="3">
        <center>
        <p><b><font style="FONT-SIZE: 12pt; FONT-FAMILY: ����,Arial; TEXT-DECORATION: none"><span style="FILTER: glow(color=Yellow,strength=4); WIDTH: 100%; COLOR: Red; LINE-HEIGHT: 40pt; POSITION: relative">��Ա�Ƚ��Խ�������ѧϰ����</span></font></b></p>
        <p>�� <a href="kaoti.asp">��ϰ�� </a></p>
        <p>�� <a href="kaoti2.asp">������ </a></p>
        </center>
      </td>
    </center>
    <td width="40%" valign="top" align="left"> 
    </td>
    <td width="40%" valign="top" align="right" style="border-right-style: solid; border-right-color: #2163FF">
      </td>
  </tr>
  <tr>
    <td width="80%" valign="top" align="center" colspan="2" style="border-right-style: solid; border-right-color: #2163FF" background="images/bg.gif">
      <table border="0" width="80%" style="font-size: 10.5pt">
        <tr>
          <td width="100%">
            <p align="center"><b><span class=f12>������</span></b> (��<%=i-1%>��)</p>
<% 
rs.movefirst
i = 1
while not rs.eof
	response.write "<p>" & i & "��" & rs("��Ŀ") & "</p>"
	response.write "<p>"
	response.write "<input onClick=record(" & i & ",1) type=radio name=Q" & i & ">" & rs("��A")
	response.write "<input onClick=record(" & i & ",2) type=radio name=Q" & i & ">" & rs("��B")
	response.write "<input onClick=record(" & i & ",3) type=radio name=Q" & i & ">" & rs("��C")
	response.write "</p>"
	rs.movenext
	i = i + 1
wend

rs2.movefirst
while not rs2.eof
	response.write "<p>" & i & "��" & rs2("��Ŀ") & "</p>"
	response.write "<p>"
	response.write " <input onClick=record(" & i & ",1) type=radio name=Q" & i & ">��ȷ"
	response.write " <input onClick=record(" & i & ",2) type=radio name=Q" & i & ">����"
	response.write "</p>"
	rs2.movenext
	i = i + 1
wend

%>    

          </td>
        </tr>
      </table>
     <form method="post" name="cf" action="cftj.asp">
      <INPUT name=fen size=0 type="hidden">
      <INPUT name=se size=20 readonly=1>


      <INPUT onclick=checkscore(1) type=button value=����><p>

     </form>
    </td>
  </tr>
  <tr>
    <td width="40%" valign="bottom" align="left"> </td>
    <td width="40%" valign="bottom" align="right" style="border-right-style: solid; border-right-color: #2163FF">
      </td>
  </tr>
</table>
<table border="0" width="777" bgcolor="#2163FF" align="left" cellspacing="0" cellpadding="0" class=f9y>
  <tr>
    <td width="100%" align="center"> <br>
      &copy; 2001&nbsp;&nbsp; ��վ��<br>
      �й��Ϸ����չɷ����޹�˾���ڷֹ�˾<br>
      ��ί���ɻ�ά�޳�����֧&nbsp; ������ά��<br>
      <br>
    </td>
  </tr>
</table>
<script language="JavaScript">showtime(0)</script>
</body>
</html>