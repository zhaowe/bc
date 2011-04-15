<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<%
Dim paraErrNo

on error resume next 

paraErrNo = Request.QueryString("ErrNo")

Set objErr = Server.CreateObject("Com_ErrorManage.clsErrorManage")
Dim iErrNo

iErrNo = objerr.ErrorDeal(Session("UserID"),paraErrNo,,,,,,,,1)
set objErr = nothing

if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If
Response.Redirect "Operate.asp"
%>