
<%
dim functionid
dim objdml
    functionid=Request.QueryString("functionid")
 'on error  resume next   
set objdml=server.CreateObject("com_usermanage.clsUserManage")
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
<title>增加子表</title>
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<div style="width: 575; height: 15"> <b><font color="blue">增加子表</font></b> </div>                                 
 <form name="addsontable" action="addfsontableout.asp" method="post">
  <table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">功能号:</font></td>
      <td width="" bgcolor="#FFFFFF"> 
        <input type=hidden name="functionid" value="<%=functionid%>">
        <%=functionid%> </td>
    </tr>
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">功能名:</font></td>
      <td width="" bgcolor="#FFFFFF"> 
        <input type=test name="functionname" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="60" align="right" bgcolor="#003333"><font color="#FFFFFF">版本名:</font></td>
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
  <input type="submit" name="button" value="提交">
<input type="reset" name="reset" value="重置"> 
  <input type="button" name="reset_button2" value="返回" onclick="self.history.back()">             
</form>             
