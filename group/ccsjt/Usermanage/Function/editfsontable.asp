
<%
dim functionid
dim dbjdml
dim objrs
dim ierrno
dim functionname

    functionid=session("functionid")
    locale=Request.QueryString("locale")
  
set objdml=server.createobject("com_usermanage1.clsFunction1")
on  error resume next
 functionname=objdml.GetFunctionName(functionid,locale)
if err.number<>0 then
    ierror=err.number
    err.clear
    response.redirect"../../Sorry.asp?ErrorNo=" & iErrNo
end if
set objdml=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ����ӱ���Ϣ</title>
<link rel="stylesheet" href="../../style.css">
</head> 


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<b><font color="blue">�޸��ӱ���Ϣ</font></b> <form name=editfunctionlocale action="editfunctionlocale.asp" method="post"> 
<input type=hidden name="functionid" value="<%=functionid %>" maxlength=50>
 <input type=hidden name="locale" value="<% =locale %>">
<table border="0" width="610" bgcolor="#000000" cellpadding="4" cellspacing="1">
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">����ID:</font></td>
    <td width="" bgcolor="#FFFFFF"><% =functionid%></td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">�汾��</font></td>
    <td width="" bgcolor="#FFFFFF"><%=locale %> </td>
  </tr>
  <tr> 
    <td width="50" align="right" bgcolor="#003333"><font color="#FFFFFF">������</font></td>
    <td width="" bgcolor="#FFFFFF"> 
      <input type=test name="functionname" value="<% =functionname%>" maxlength=50>
    </td>
  </tr>
</table> 
<br>
<input type=submit name=button value="�ύ">
<input type=reset name=button2 value="����">                                                     
<input type="button" name="reset_button2" value="ȡ��" onclick="self.history.back()">