<!--#include file="public.inc"-->

<%
dim groupid
dim objdml
    groupid=Request.QueryString("groupid") 
 on error  resume next   
set objdml=server.CreateObject("com_usermanage1.clsusermanage1")
set allcola=server.CreateObject("adodb.recordset")
set allcola=objdml.GetLocale()
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing	
%>
    
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����ӱ�</title>
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<script language="vbscript" >     
<!--     
sub datacheck()       
 if addgroup.groupname.value="" then     
    msgbox "������������Ϊ�� ",64,"��ע��!"     
    focusto(0)     
 exit sub     
 end if      
  addgroup.submit     
 end sub      
-->     
</script>
<div style="width: 575; height: 43"> <b><font color="blue">�����ӱ�</font></b> <br>
</div>                                 
 <form name="addsontable" action="addsontable.asp" method="post">
  <table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">���:</font></td>
      <td width="" bgcolor="#FFFFFF"> 
        <input type=hidden name="groupid" value="<%=groupid%>" >
        <%=groupid%> </td>
    </tr>
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">����:</font></td>
      <td width="" bgcolor="#FFFFFF"> 
        <input type=test name="groupname" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">�汾��:</font></td>
      <td bgcolor="#FFFFFF"> 
        <select name="locale" size=1 >
          <% 
       do while not allcola.eof 
     %> 
          <option value="<%=allcola(0)%>"><%=allcola(1)%></option>
          <% allcola.movenext%><% loop %> 
        </select>
        <% allcola.close %></td>
    </tr>
  </table>                                    
  <br>
  <input type="submit" name="button" value="�ύ">
<input type="reset" name="reset" value="����"> 
  <input type="button" name="reset_button2" value="����" onclick="self.history.back()">             
             
</form>             
