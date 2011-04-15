<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="TypeConvert.asp"-->
<%
Dim paraErrNo
Dim paraLocale
Dim paraProtocol

paraErrNo = Request.QueryString("ErrNo")
paraLocale = Request.QueryString("Locale")
paraProtocol = Request.QueryString("Protocol")

Dim intLocale
Dim intProtocol

intLocale=LocaleStrToInt(paraLocale)
intProtocol=ProtocolStrToInt(paraProtocol)

Dim objErr
Set objErr = Server.CreateObject("Com_ErrorManage.clsErrorManage")
Dim iErrNo

on error resume next
iErrNo = objerr.LocaleTypeDeal(paraErrNo,intLocale,intProtocol,,,1)
set objErr = nothing
if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

Response.Redirect "EditError.asp?ErrNo=" & paraErrNo
%>