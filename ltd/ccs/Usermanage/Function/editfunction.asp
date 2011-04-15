<!--#include file="public.inc"-->
<%
dim functionid
dim objdml
dim rs
dim ierror
    functionid=Request.QueryString ("which")
    session("functionid")=functionid
on error resume next
set objdml=server.CreateObject("com_UserManage.clsFunction")
set rs=server.CreateObject("adodb.recordset")
set rs=objdml.getfunctioninfo(functionid)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改组信息</title>
<link rel="stylesheet" href="../../style.css">
</head> 


<body bgcolor="#FFFFFF">
<p><b><font color="blue">修改功能信息</font></b></p>
<form name="editfunction" action="editfunctionout.asp" method="post">
  <table border="0" width="610" bgcolor="#000000" cellpadding="4" cellspacing="1">
    <tr> 
      <td width="119" align="right" bgcolor="#003333"><font color="#FFFFFF">功能名:</font></td>
      <td width="481" bgcolor="#FFFFFF"> 
        <input type=test name="functionname" value="<% Response.Write rs("functionname")%>" maxlength=50>
      </td>
    </tr>
    <tr>  <td width="119" align="right" bgcolor="#003333"><font color="#FFFFFF">父功能:</font></td>
      <td width="481" bgcolor="#FFFFFF"> 
     <select name="fFunctionid" size="1" ><option value=""></option>
          <% 
          dim funinfo
                 set funinfo=server.CreateObject ("ADODB.Recordset")
                 set funinfo=objdml.getallfunction("zh")
             if err.number<>0 then
                ierror=err.number
                err.clear
                set objdml=nothing
                response.redirect "../../Sorry.asp?Errorno="&ierror
              end if
              funinfo.Movefirst
              do while not funinfo.EOF
            
              %> 
          <option value="<%=funinfo("Functionid")%>"<% 
               if funinfo("Functionid")=rs("fFunctionid") then%>selected<%End IF%>> 
          <%=funinfo("FunctionName")%></option>
          <%funinfo.MoveNext 
            loop
           %> 
        </select></td>
        </tr>
      <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">功能类型:</font></td>
      <td bgcolor="#FFFFFF"> 
      <%dim f
      f=trim(rs("functiontype"))
  
      %>
        <input type="checkbox" name="functiontypem" <%if instr(f,"M") then%>checked <%end if%>>菜单
        <input type="checkbox"  name="functiontypef" <%if instr(f,"F")then%> checked<%end if%> >功能
        <input type="checkbox" name="functiontypep" <%if instr(f,"P") then%>checked<%end if %>>页面
      </td>
    </tr>
    <tr> 
      <td width="119" align="right" bgcolor="#003333"><font color="#FFFFFF">序号:</font></td>
      <td width="481" bgcolor="#FFFFFF"><% Response.Write rs("OrderNum")%></td>
    </tr>
    <tr> 
      <td width="119" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
      <td width="481" bgcolor="#FFFFFF"><% Response.Write rs("description")%></td>
    </tr>
  </table> 
<br>
  <table width="610" border="0">
    <tr>
      <td width="229"> 
        <input type="hidden" name="functionid" value="<% Response.write functionid%>" maxlength=50>
        <input type=submit name="submit" value=提交>
        <input type=reset name="reset" value=重置>
        <input type="button" name="reset_button2" value="返回" onClick="self.history.back()">
      </td>
      <td align="center" width="371">[ <a href="<%="fsontablerun.asp"&"?functionid="& functionid%>">子表管理</a> 
        ]|[ <a href="functioninfo.asp">返回功能管理页</a> ] </td>
    </tr>
  </table>
</form>  
<%set objdml=nothing%>                         
