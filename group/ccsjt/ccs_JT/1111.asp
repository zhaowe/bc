<html>
<head>
<title>������������������</title>
</head>

<body>
<form name="myform" method="post">
<input type=hidden name=valee value="12121212121212">
<select  name="depar"  onChange="javascript:changeclass2()" onfocus="javascript:changeclass2()"> 
<option value="ȫ��" selected>���в���</option>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("OledbStr")
Sql="Select * From CompanyLocale"
'Sql="Select distinct depar From cwys_km"
Set Rs=Server.CreateObject("ADODB.RecordSet")
Rs.Open Sql,Conn,1,1
do while not rs.eof
'Response.write "<option value='"&rs("CompanyName")&"'>"&rs("CompanyName")&"</option>"
'Response.write "<option value='"&rs("depar")&"'>"&rs("depar")&"</option>"
Response.write "<option value='"&trim(rs("CompanyName"))&"'>"&rs("CompanyName")&"</option>"
rs.movenext
loop
Response.write "</select>"
Rs.Close
Set Rs = Nothing
%>


<select  name="kemu"> 
  <option value="ȫ��" selected>���п�Ŀ</option>
</select>
</form> 



<%
Sql="Select * From cwys_km"
Set Rs=Server.CreateObject("ADODB.RecordSet")
Rs.Open Sql,Conn,1,1

do while not rs.eof
str=str&trim(rs("depar"))&"/"&trim(rs("fkmcode"))&"/"&trim(rs("fkmshuom"))&"," 
rs.movenext
loop
Rs.Close
Set Rs = Nothing
%>
<%
i=3245683.336
c=cdbl(i)
%>
<%=c%>

<!--������ʵ�ֶ�̬�ı���һ���˵��Ľű�����--> 
<script  LANGUAGE="javascript"> 
arr="<%=str%>".split(","); 
a=arr.length 
ar=new Array() 
for (i=0;i<a;i++)
{ 
 ar[i]=arr[i].split("/"); 
} 

function  changeclass2() {
 document.myform.kemu.length=1
 lid=myform.depar.value;  
 for  (i=0;i<a;i++)  {
   if  (ar[i][0]  ==  lid) {
  document.myform.kemu.options.add(new Option(ar[i][2],ar[i][1])); 
   }
 }
}
</script>
<p>
<%=now%>
</p>
<%
a=rnd(25)
%>
<%=a%>
</body>
</html>