<!--#include file="public.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>增加功能信息</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<script language="vbscript">  
<!--  
sub datacheck()  
 if addfunction.description.value="" then  
 msgbox "描述不能为空 ",64,"请注意!"  
 exit sub   
 end if
 if addfunction.functionname.value="" then  
 msgbox "功能名不能为空 ",64,"请注意!"  
 exit sub   
 end if
  addfunction.submit
end sub   
-->  
</script>
<b><font color="#0000FF">增加功能信息</font></b> <br>
<form name="addfunction" action="addfunout.asp" method="post">
  <table border="0" width="610" cellpadding="4" cellspacing="1" bgcolor="#000000">
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">功能名:</font></td>
      <td bgcolor="#FFFFFF" > 
        <input type=test name="functionname" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
      <td bgcolor="#FFFFFF"> 
        <input type=test name="description" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">父功能:</font></td>
      <td bgcolor="#FFFFFF"> 
        <select name="ffunctionid" size="1" ><option value=""></option>
          <% 
          dim funinfo
          dim objdml
                 set objdml=CreateObject("com_usermanage1.clsFunction1")
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
          <option value="<%=funinfo("functionid")%>"> 
          <%=funinfo("functionName")%></option>
          <%funinfo.MoveNext 
            loop
           %> 
           
        </select>
      </td>
    </tr>
    <tr> 
      <td width="130" align="right" bgcolor="#003333"><font color="#FFFFFF">功能类型:</font></td>
      <td bgcolor="#FFFFFF"> 
        <input type="checkbox" name="functiontype1" value="M" >菜单
        <input type="checkbox" name="functiontype2" value="F">功能
        <input type="checkbox" name="functiontype3" value="P">页面
      </td>
    </tr>
  </table>
  <br>
  <input type="button" name="button" value="提交" onclick="datacheck">
  <input type="reset" name="reset_button" value="重填">
  <input type="button" name="reset_button2" value="返回" onclick="self.history.back()">
</form>                              
