<!--#include file="public.inc"-->

<%
dim functionid
dim conflict
dim fun
    functionid=Request.QueryString ("which")
dim objdml
dim functioninfo
set objdml=server.createobject("com_usermanage1.clsFunction1")
on error resume next
    functioninfo=objdml.GetFunctionInfo(functionid,locale)
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
<title>ɾ��������Ϣ</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<b><font color="blue">ɾ��������Ϣ</font></b> <br>
<a href="functioninfo.asp"> </a><br>
ɾ������ID:<font face="����_GB2312" size="3"><%=functionid%></font> 
<form name=delfunction action="delfunctionout.asp" method="post"> 
  <input type=hidden name=functionid value="<% =functioninfo("functionid")%>">
  <table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">������:</font></td>
      <td width="369" bgcolor="#FFFFFF"><%=functioninfo("functionname")%></td>
    </tr>
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">����:</font></td>
      <td width="369" bgcolor="#FFFFFF"><%=functioninfo("description")%></td>
    </tr>
  </table>       
  <table width="610" border="0" cellpadding="4" cellspacing="1">
    <tr>
      <td width="222"><% set functioninfo=nothing %> 
        <input type=submit name="submit" value=���ɾ��>
        <!--
  <input type=reset name="reset_button" value=����ɾ��>                             
  --> </td>
      <td width="367" align="center"><a href="functioninfo.asp">���ع��ܹ�����ҳ </a> 
        &nbsp;&nbsp;&nbsp;&nbsp; <font face="����_GB2312" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        </font></td>
    </tr>
  </table>
</form>   
 </body>                                   
