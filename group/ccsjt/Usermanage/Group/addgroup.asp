<!--#include file="public.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="style.css">
<title>增加组</title>
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<script language="vbscript" >     
<!--     
sub datacheck()
 if addgroup.groupname.value="" then     
    msgbox "“中文组名”不能为空 ",64,"请注意!"      
	exit sub
 End if
 if addgroup.description.value="" then     
    msgbox "“描述”不能为空 ",64,"请注意!"
	exit sub
 end if
  addgroup.submit     
 end sub      
-->     
</script>
<b><font color="blue">增加组</font> <br>
</b> 
<form name="addgroup" action="addgpout.asp" method="post">
  <table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr> 
      <td width="78" align="right" bgcolor="#003333"><font color="#FFFFFF">中文组名:</font></td>
      <td width="511" bgcolor="#FFFFFF"> 
        <input type=test name="groupname" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="78" align="right" bgcolor="#003333"><font color="#FFFFFF">描述:</font></td>
      <td width="511" bgcolor="#FFFFFF"> 
        <input type=test name="description" value="" maxlength=50>
      </td>
    </tr>
  </table>  
<br>                                  
<input type="button" name="button" value="提交" onclick="datacheck">
<input type="reset" name="reset" value="重置">
  <input type="button" name="reset_button2" value="返回" onclick="self.history.back()">             
              
</form>             
