
<%
dim groupid
dim objdml
dim localeinfo
   groupid=Request.QueryString("Which")
set objdml=Server.CreateObject("Com_UserManage1.clsUserManage1")
on error resume next
set localeinfo=Server.CreateObject("ADODB.Recordset")
set  localeinfo=objdml.GetGroupLocale(groupid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
 session("groupid")=groupid
 
howmanyfield=localeinfo.fields.count-1

%>

<html> 
<head>
<title>���ӱ����</title>
<link rel="stylesheet" href="../../style.css">
</head>

<body bgcolor="#FFFFFF">
<font color=blue><strong>���ӱ����</strong></font><br>
<br>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {
	form1.text.value="edit"
	form1.submit()
	return true 
}

function button2_onclick() {
	form1.text.value="del"
	form1.submit()
	return true 
}

//-->
</SCRIPT>
<FORM action="bosom1.asp" method="post" name="form1">
  <table border=0 cellPadding=4  cellSpacing=1 width="610" bgcolor="#000000">
    <TR bgcolor="#003333" Valign=top align=center> 
      <td><font color="#FFFFFF">ѡ��</font></td>
      <TD><font color="#FFFFFF">�汾</font></TD>
      <TD><font color="#FFFFFF">����</font></TD>
    </TR>
    <% do while not localeinfo.eof%> 
    <tr> 
      <td align=top bgcolor="#FFFFFF" > 
        <INPUT id=radio1 name=locale type=radio value="<%=localeinfo(1)%>" <%if localeinfo(1="zh") then  Response.Write "checked" end if%> >
      </td>
      <% for j=1 to howmanyfield%> 
      <td valign=top bgcolor="#FFFFFF" ><%=localeinfo(j)%></td>
      <%  next  %> </tr>
    <% 
localeinfo.movenext
loop
localeinfi.close  
set objdml=nothing
%> 
  </table>
  <table width="610" border="0" cellpadding="4" cellspacing="1">
    <tr>
      <td width="258"> 
        <input name=text type=hidden>
        <input name=button1 type=button value=�޸� language=javascript onClick="return button1_onclick()">
        <input name=button2 type=button value=ɾ�� language=javascript onClick="return button2_onclick()">
      </td>
      <td width="331" align="center"> <% my_link="addsuntable.asp" &"?groupid="&groupid %> 
        [ <a href="<%=my_link%>">������汾</a> ][ <a href="groupinfo.asp">���������</a> 
        ]</td>
    </tr>
  </table>
</FORM>