<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<%
Dim paraOperate
Dim paraErrNo
Dim paraReasonIn
Dim paraSolution
Dim paraClassAType
Dim paraClassBType
Dim paraPrgName
Dim paraErrGoto
Dim paraErrType
Dim paraLocale
Dim paraProtocol
Dim paraNameOut
Dim paraSolutionOut

paraOperate = Request.Form("hidOperate")

'Error表相关值
paraErrNo = Request.Form("hidErrNo")  
paraReasonIn = Request.Form("txtReasonIn")
paraSolution =  Request.Form("txtSolution")
paraClassAType = Request.Form("selClassA")
paraClassBType = Request.Form("selClassB")
paraPrgName = Request.Form("txtErrPrg")
paraErrGoto = Request.Form("txtErrGoto")
paraErrType = Request.Form("selType")
'Locale表相关值
paraLocale = 2 'zh
paraProtocol = 1 'http
paraNameOut = Request.Form("txtNameOut")
paraSolutionOut = Request.Form("txtSolutionOut")

Set objErr = Server.CreateObject("Com_ErrorManage.clsErrorManage")
Dim iErrNo
Dim iLocale
Dim iOperate
on error resume next
'**********************************************
'*添加纪录的操作，ErrorNo值来自ErrNoBack返回值*
'*修改纪录的操作，ErrorNo值由上一页面传来     *
'**********************************************
If paraOperate = "Add" Then
	iOperate = 2
	iErrNo = objErr.ErrorDeal(Session("UserID"),,paraReasonIn,paraSolution,paraClassAType,paraClassBType,paraPrgName,paraErrGoto,paraErrType,iOperate,ErrNoBack)
	if Err.number <> 0 then
		iErrNo = Err.number
		Err.Clear
		Set objErr = Nothing
		Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
	End If
	Dim ErrNoBack
	iLocale =  objErr.LocaleTypeDeal(ErrNoBack,paraLocale,paraProtocol,paraNameOut,paraSolutionOut,iOperate)
ElseIf paraOperate = "Edit" Then
	iOperate = 4
	iErrNo = objErr.ErrorDeal(Session("UserID"),paraErrNo,paraReasonIn,paraSolution,paraClassAType,paraClassBType,paraPrgName,paraErrGoto,paraErrType,iOperate)
	
	if Err.number <> 0 then
		iErrNo = Err.number
		Err.Clear
		Set objErr = Nothing
		Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
	End If
	iLocale =  objErr.LocaleTypeDeal(paraErrNo,,,paraNameOut,paraSolutionOut,iOperate)
End If
Set objErr = Nothing
if Err.number <> 0 then
	iLocale = Err.number
	Err.Clear
	Response.Redirect "Sorry.asp?ErrNo=" & iLocale
End If

If paraOperate = "Add" Then
	Response.Redirect "EditError.asp?ErrNo=" & ErrNoBack
elseif paraOperate = "Edit" then
	Response.Redirect "EditError.asp?ErrNo=" & paraErrNo
End IF
%>