

<%
dim GroupId
dim GroupName
dim locale
dim objdml
dim rs
dim ierrno
   groupid=Request.Form("GroupID")
   locale=Request.Form("Locale")
   groupname=Request.Form("GroupName")
set objdml=server.CreateObject("com_usermanage1.clsusermanage1")
on error resume next
 ierrno=objdml.AddGroupLocale(groupid,groupname,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
	set objdml=nothing	
End If    
set objdml=nothing
Response.Redirect "gsontablerun.asp"& "?Which="& groupid
%>
