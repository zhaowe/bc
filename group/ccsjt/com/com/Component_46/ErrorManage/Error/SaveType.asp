<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="TypeConvert.asp"-->
<%
Dim paraOperate
Dim paraErrNo
Dim paraLocale
Dim paraProtocol
Dim paraNameOut
Dim paraSolutionOut

paraOperate = Request.Form("hidOperate")
'*****************************
'   判断是添加还是修改记录
'*****************************

if paraOperate="Edit" then
	TempLocale = Request.Form("hidLocale")
	TempProtocol = Request.Form("hidProtocol")
	paraLocale = LocaleStrToInt(TempLocale)
	paraProtocol = ProtocolStrToInt(TempProtocol)
elseif paraOperate="Add" then
	paraLocale = Request.Form("selLocale")
	paraProtocol = Request.Form("selProtocol")
end if

paraErrNo = Request.Form("hidErrNo")
paraNameOut = Request.Form("txtNameOut")
paraSolutionOut = Request.Form("txtSolution")

Set objErr = Server.CreateObject("Com_ErrorManage.clsErrorManage")
Dim iErrNo
Dim iOperate
on error resume next

If paraOperate = "Add" Then
	iOperate = 2
	iErrNo = objErr.LocaleTypeDeal(paraErrNo,paraLocale,paraProtocol,paraNameOut,paraSolutionOut,iOperate)
ElseIf paraOperate = "Edit" Then
	iOperate = 4
	iErrNo =  objErr.LocaleTypeDeal(paraErrNo,paraLocale,paraProtocol,paraNameOut,paraSolutionOut,iOperate)
End If
	
Set objErr = Nothing
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

Response.Redirect "EditError.asp?ErrNo=" & paraErrNo
%>