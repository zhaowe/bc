<%@ Language=VBScript %>
<!--#include file="../../include/UserCheck.asp"-->

<%
AgentOffice=Request.form("AgentOffice")
AgentType=Request.Form("AgentType")
AgentName=Request.Form("AgentName")
AgentShortName=Request.Form("AgentShortName")
lxrAdd=Request.Form("lxrAdd")
lxrName=Request.Form("lxrName")
lxrPhone=Request.Form("lxrPhone")
AgentCity=Request.Form("AgentCity")
OpenBank=Request.Form("OpenBank")
OpenAccount=Request.Form("OpenAccount")
ProtocalNo=Request.Form("ProtocalNo")
ProtocalDate=Request.Form("ProtocalDate")
AgentEntity=Request.Form("AgentEntity")
FAgentID=Request.Form("FAgentID")  '上级代理点AgentID
Operator=session("UID")
UseObject=application("UseObject")


set objAgent=server.CreateObject("com_Agent.clsAgent")
objAgent.UseObject = application("UseObject")
objAgent.locale = application("locale")
objAgent.protocol = application("protocol")
   
on error  resume next

ret=objAgent.NewAgent(AgentOffice, AgentType, _
AgentName, AgentShortName, UseObject, AgentCity, lxrAdd, _
lxrPhone, lxrName, OpenBank, OpenAccount, ProtocalNo, _
ProtocalDate, Operator, AgentEntity,FAgentID)

 If Err.Number<>0 then
   ErrNo=Err.Number
   Err.clear
   set objAgent=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
 End if
 set objAgent=nothing

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<p><b><font color="#0000FF">恭喜！新增加盟点成功</font></b></p>
<p>[ <a href="viewAgentList.asp">返回</a> ]</p>
</BODY>
</HTML>
