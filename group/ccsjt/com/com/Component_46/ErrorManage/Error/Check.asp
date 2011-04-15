<!--#include file="TypeConvert.asp"-->
<!--#include file="DbClass.asp"-->
<%
Dim paraUser
Dim paraPwd
paraUser = Request.Form("txtUser")
paraPwd = Request.Form("txtPwd")

Dim objDML
Dim objRs
Dim strSql
Dim iErrNo

on error resume next

Set objDML = Server.CreateObject("Com_DML.ClsDML")

strSql = "Select * from ErrUser Where UserName='" & paraUser & "' and UserPwd='" & paraPwd & "'"
Set objRs = Server.CreateObject("Adodb.Recordset")
Set objRs = objDML.ExeSelect(strSql,DbClass)
set objDML = nothing

if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

If not objRs.EOF then
	Session("UserID") = objRs("UserId")
	Session("LoginOk") = True
	Response.Redirect "Operate.asp"
Else
	iErrNo = 10052 '登录用户名或密码错误
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If
%>