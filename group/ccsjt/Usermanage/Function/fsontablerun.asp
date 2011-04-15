
<!--#include file="public.inc"-->
<%
dim functionid
dim objdml
dim localeinfo
   functionid=Request.QueryString("functionid")
set objdml=server.createobject("com_usermanage1.clsFunction1")
on error resume next
set localeinfo=server.CreateObject("adodb.recordset")
set  localeinfo=objdml.GetFunctionLocale(functionid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
 session("functionid")=functionid
howmanyfield=localeinfo.fields.count-1
%>

<html> 
<head>
<title>功能子表管理</title>
<link rel="stylesheet" href="../../style.css">
</head>

<body bgcolor="#FFFFFF">
<font color=blue><strong>功能子表管理</strong></font><br>
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
    <tr valign=top align=center bgcolor="#003333"> 
      <td width="37"><font color="#FFFFFF"></font></td><!--
      <td width="261"><font color="#FFFFFF">功能号</font></td>-->
      <td width="56"><font color="#FFFFFF">版本</font></td>
      <td width="215"><font color="#FFFFFF">功能名</font></td>
    </tr>
    <% do while not localeinfo.eof%> 
    <tr> 
      <td align=top bgcolor="#FFFFFF" width="37" > 
        <INPUT id=radio1 name=locale type=radio value="<%=localeinfo(1)%>" <%if i like"zh" then Response.Write  "checked" end if%>>
      </td>
      <% for j=1 to howmanyfield%> 
      <td valign=top bgcolor="#FFFFFF"><%=localeinfo(j)%></td>
      <%  next  %> </tr>
    <% 
localeinfo.movenext
loop
localeinfo.close  
set objdml=nothing
%> 
    </table>
  <br>
  <table width="610" border="0">
    <tr>
      <td width="219"> 
        <input name=text type=hidden>
        <input name=button1 type=button value=修改 language=javascript onClick="return button1_onclick()">
        <input name=button2 type=button value=删除 language=javascript onClick="return button2_onclick()">
        <input name=button22 type=button value=返回 onClick="self.history.back()">
      </td>
      <td align="center" width="381"><% my_link="addfsontable.asp" &"?functionid="&functionid %> 
        [ <a href="functioninfo.asp">返回功能管理页</a> ][ <a href="<%=my_link%>">增加功能版本</a> ]</td>
    </tr>
  </table>
</FORM>