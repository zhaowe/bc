<%@ Language=VBScript %>
<% 
'������ʼ����
session("OpenWin")="Y" 

Session("UID")=""
Session("LoginID")=""
Session("AgentID")=""
Session("IntraLoginOk")=False


Response.Redirect ("../public/public.asp")

%>
