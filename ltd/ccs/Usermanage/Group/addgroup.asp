<!--#include file="public.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="style.css">
<title>������</title>
</head>


<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<script language="vbscript" >     
<!--     
sub datacheck()
 if addgroup.groupname.value="" then     
    msgbox "����������������Ϊ�� ",64,"��ע��!"      
	exit sub
 End if
 if addgroup.description.value="" then     
    msgbox "������������Ϊ�� ",64,"��ע��!"
	exit sub
 end if
  addgroup.submit     
 end sub      
-->     
</script>
<b><font color="blue">������</font> <br>
</b> 
<form name="addgroup" action="addgpout.asp" method="post">
  <table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr> 
      <td width="78" align="right" bgcolor="#003333"><font color="#FFFFFF">��������:</font></td>
      <td width="511" bgcolor="#FFFFFF"> 
        <input type=test name="groupname" value="" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="78" align="right" bgcolor="#003333"><font color="#FFFFFF">����:</font></td>
      <td width="511" bgcolor="#FFFFFF"> 
        <input type=test name="description" value="" maxlength=50>
      </td>
    </tr>
  </table>  
<br>                                  
<input type="button" name="button" value="�ύ" onclick="datacheck">
<input type="reset" name="reset" value="����">
  <input type="button" name="reset_button2" value="����" onclick="self.history.back()">             
              
</form>             
