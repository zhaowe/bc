����<%

 Set conn = Server.CreateObject("ADODB.Connection")
 conn.Open "provider=sqloledb;server=10.254.0.41;database=NorthWind;uid=sa;pwd=szx6275;" 
 
Set Rs=server.CreateObject ("ADODB.Recordset")
Rs.LockType=3
Rs.CursorType=3
set Rs.activeConnection=conn

sql="Select employeeid,lastname from employees order by employeeid"
rs.Source =sql
 rs.Open 
    %>

��<html>
����<head>
����<title>��ˢ��ҳ���ѯ�ķ���</title>
����<meta http-equiv="Content-Type" content="text/html" charset="gb2312">
����</head>

����<script language="javascript">
����
����function search_onclick()
����{
����//�õ�ɸѡ��Ա������
����searchtext=window.searchContent.value;
//�����Ƴ������в�ѯ����б��е�ѡ��
����j=searchObj.length;
����for(i=j-1;i>=0;i--)
����{
����searchObj.remove(i);
����}
����if(searchtext!=""){
����//��ʾ����ɸѡ�����Ĺ�Ա
����j=searchSource.length;
����for(i=0;i<j;i++){
����searchsource=searchSource.options(i).text;
����k=searchsource.indexOf(searchtext);
����if(k!=-1){
����option1=document.createElement("option");
����option1.text=searchsource;
����option1.value=searchSource.options(i).value;
����searchObj.add(option1);
����}
����}
����}
����else{
����//���û�������ѯ��������ʾ���й�Ա
����j=searchSource.length;
����for(i=0;i<j;i++){
����searchsource=searchSource.options(i).text;
����option1=document.createElement("option");
����option1.text=searchsource;
����option1.value=searchSource.options(i).value;
����searchObj.add(option1);
����}
����}
����}
����
����</script>

����<body bgcolor="#FFFFFF" text="#000000">


<table width="80%" border="1">
����<tr>
����<td>
����<input type="text" name="searchContent">
����<input type="button" name="Button" value="�顡��ѯ "  onclick="javascript:return search_onclick()">
����</td>
����</tr>
����<tr>
����<td>��ѯ���<br>
����<select name="searchObj" size="20">
����
����<%while not rs.EOF%>
����<option value="<%=rs(0)%>"><%=rs(1) %></option>
    <%
     rs.MoveNext 
     wend
    %>
����</select>
����
<select name="searchSource" size="10"  style="display:none">
<%rs.MoveFirst 

do while  not rs.EOF %>

<option value="<%=rs(0)%>"><%=rs(1)%></option>
����<%
rs.MoveNext 
loop
%>
</select>
</td>
</tr>
</table>

����</body>
����</html>����

















