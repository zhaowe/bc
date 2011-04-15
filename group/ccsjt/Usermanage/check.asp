<%
dim loginid
dim password
dim objRs
Dim ObjUser 
dim LoginPass
Dim iErrNo
dim FuncStr


session("LoginID")=""
LoginID=trim(request.form("loginid"))
Password=trim(request.form("password"))
Set ObjUser = Server.CreateObject("Com_UserManage1.clsUserManage1")

on error resume next
Set objRs = Server.CreateObject("Adodb.Recordset")
LoginPass = ObjUser.CheckLoginID(LoginID,Password,application("UseObject"))

if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	set ObjUser=nothing
	Response.Redirect "../sorry/Sorry.asp?ErrorNo=" & iErrNo
End If
set objRs = ObjUser.GetLoginInfo(LoginID,application("UseObject"))
if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	set ObjUser=nothing
	Response.Redirect "../sorry/Sorry.asp?ErrorNo=" & iErrNo 
End If
Session("UID")=ObjRs("UserID")

'if objRs.EOF and objRs.BOF then
'   Response.Write "aaaaaaaa"
'  else
'   Response.Write "ooooooo"
'   Response.Write(ObjRs("UserID"))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(1))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(2))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(3))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(4))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(5))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(6))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(7))
'   Response.Write "ooooooo"
'   Response.Write(ObjRs(8))
'   
'end if
 
Session("LoginID")=LoginID
Session("AgentID")=ObjRs("AgentID")
Session("IntraLoginOk")=true
FuncStr=ObjUser.GetFuncStr(ObjRs("UserID"))
Session("FuncStr")=FuncStr
set ObjRs=nothing
set ObjUser=nothing
session("IsOutAddrPass")=true
Response.Redirect "../index.asp"
%>