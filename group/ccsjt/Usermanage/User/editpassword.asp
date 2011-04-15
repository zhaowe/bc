<%@ Language=VBScript %>
<%UserId=Request.QueryString("UserID")
%>
<HTML>
<HEAD>
<script ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
sub btnQuery_onclick
dim name
dim loginid
dim contactinfo
dim password1
dim password2
dim password
cnstr=""
titlestr="[您请注意]"+chr(13)+chr(13)
errstr=""
password1=trim(document.editpassword.password1.value)
if(password1="")then
cnt=cnt+1
cntstr=cstr(cnt) 
errstr=errstr+cntstr+"."+"请输入您的新密码"+chr(13)
else
document.editpassword.password1.value=password1
end if
password2=trim(document.editpassword.password2.value)
if(password2="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请输入您的校验密码"+chr(13)
else
document.editpassword.password2.value=password2
end if
if password1<>password2 then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"您两次输入的密码不同，请再试！"+chr(13)
end if
if len(password1)>=10 then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请您输入一个小于10位的密码"+chr(13)
end if
if cnt<>0 then
alert(errstr)
else
editpassword.submit()
end if
end sub

-->
</script>
<link rel="stylesheet" href="../../style.css">
</HEAD>
<BODY>
<form name=editpassword Method="post" action="editpassword_result.asp?UserID=<%=UserID%>">
  <p><b><font color="#0000FF">修改用户密码</font></b></p>

    
  <table width="610" border="0" bgcolor="#000000" cellspacing="1" cellpadding="4">
    <tr> 
      <td width="25%" height="26" bgcolor="#003333"> 
        <div align="right"><font color="#FFFFFF">请输入您的新用户密码 ：</font></div>
      </td>
      <td width="75%" height="26" bgcolor="#FFFFFF"> 
        <input type=password name=password1 >
      </td>
    </tr>
    <tr> 
      <td width="25%" bgcolor="#003333"> 
        <div align="right"><font color="#FFFFFF">请校验您的新密码：</font></div>
      </td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input type=password name=password2 >
      </td>
    </tr>
  </table>
  <p align="left"> 
    <input type=button name=btnQuery id=btnQuery value="修改密码">
    <input type=button id=btnQuery1 value="返回" onclick="self.history.back()">
  </p>
</form>
</BODY>
</HTML>
