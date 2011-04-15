<%@ Language=VBScript %>
<% 
'设置起始参数
session("OpenWin")="Y" 

Session("UID")=""
Session("LoginID")=""
Session("AgentID")=""
Session("IntraLoginOk")=False


Response.Redirect ("../public/public.asp")

%>
