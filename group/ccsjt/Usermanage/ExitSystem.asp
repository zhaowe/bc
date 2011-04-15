<%@ Language=VBScript %>
<%
Session("IntraLoginOK")=false
Session("UID")=""
Session("AgentID")=""
Session("FuncStr")=""
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	document.form1.submit()
	return true
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM action="../UserManage/login.asp" method=POST id=form1 name=form1 target="_top">
<P>&nbsp;</P>
</FORM>
</BODY>
</HTML>
