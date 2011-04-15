<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0041)http://szxair/managesystem/jane/caiwu.htm -->
<html>
<head>

<title>预算控制系统</title>
<meta content="text/html; charset=gb2312" http-equiv="Content-Type">
<meta content="MSHTML 5.00.2614.3500" name="GENERATOR">
<meta content="FrontPage.Editor.Document" name="ProgId">


<script language="javascript">
function send_onsubmit() {
	if (login.loginid.value ==""){
		alert("请输入真实用户，,谢谢合作")
	return false
		}

	if (login.password.value==""){
		alert("“密码不能为空”,请输入你的密码!谢谢!,谢谢合作")
	return false
		}
}
</script>



</head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" background="image/fly1.jpg" rightMargin=0>
<script ID="clientEventHandlersJS" LANGUAGE="javascript"><!--
function HzgOK_onclick() {
	
	document.UserSelectFormDisp.submit()
}
function ChangePasswordOK_onclick() {
	
	document.ChangePassword.submit()
}
function HzgReturn_onclick() {
	document.location.href="../"
}
//--></script>
<br>
<br>
<p align=center><font size=6 color=red>欢迎进入财务预算系统登陆页面</font></P>
 <table border="0"  cellPadding="0" cellSpacing="0">
        <tbody>
        <tr>
          <td height="60" vAlign="top" width=1000>
            <br>
            <table align="center" border="0" cellspacing="0" cellpadding="0" width=100%>
            
                 <form action="checkgz.asp" method="post" name="login" onsubmit="return send_onsubmit()">
                 <tr height=70> 
                     <td align="middle">
                     </td>
                 </tr>             
                 <tr> 
                     <td align="middle" ><font face="Arial, Helvetica, sans-serif">用户</font>
                     <input name="loginid" size="16" maxlength="20" style="WIDTH: 91px; HEIGHT: 21px">
                     </td>
                 </tr>
                 <tr> 
                     <td align="middle"><font face="Arial, Helvetica, sans-serif">密码</font>
                        <input name="password" type="password" size="16" maxlength="10" style="WIDTH: 91px; HEIGHT: 21px">
                     </td>
                 </tr>
                 <tr>
                     <td align="middle">&nbsp;&nbsp;&nbsp;
                        <input type="submit" value="登录" name="Submit" style="WIDTH: 40px; HEIGHT: 25px">
                        <input type="reset" value="重填" name="reset" style="WIDTH: 40px; HEIGHT: 25px">
                     </td>
                 </tr>
                 </form>
           </table>
</body>
</HTML>
