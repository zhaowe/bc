
<%
dim groupid
dim groupinfo
dim objdml
dim ierrno
dim localeinfo
dim i
dim howmanyfield
    locale=Request.QueryString ("locale")
    groupid=session("groupid")
     session.Abandon
    
set objdml=server.createobject("com_usermanage.clsusermanage")

on error resume next
 str=objdml.DelGroupLocale(groupid,locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
Response.Redirect"gsontablerun.asp?"&"which="&groupid
%>
